#!/usr/bin/env bash
#
# Install the CV kit into a career-ops checkout.
#
#   ./install.sh /path/to/career-ops
#   ./install.sh --dry-run /path/to/career-ops
#
# Copies five CV templates, the Word (.docx) builder, and the cv-word skill.
# Nothing career-ops ships is overwritten, so `node update-system.mjs apply`
# keeps working. Re-run this after an update to restore anything it removed.

set -euo pipefail

DRY_RUN=0
NO_NPM=0
TARGET=""

for arg in "$@"; do
	case "$arg" in
		--dry-run) DRY_RUN=1 ;;
		--no-npm) NO_NPM=1 ;;
		-h|--help)
			sed -n '2,10p' "$0" | sed 's/^# \{0,1\}//'
			exit 0
			;;
		-*)
			echo "Unknown option: $arg" >&2
			exit 2
			;;
		*) TARGET="$arg" ;;
	esac
done

KIT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Fall back to a sibling career-ops checkout, the layout most people end up with.
if [ -z "$TARGET" ] && [ -d "$KIT_DIR/../career-ops" ]; then
	TARGET="$KIT_DIR/../career-ops"
fi

if [ -z "$TARGET" ]; then
	echo "Usage: ./install.sh [--dry-run] [--no-npm] /path/to/career-ops" >&2
	exit 2
fi

if [ ! -d "$TARGET" ]; then
	echo "Not a directory: $TARGET" >&2
	exit 1
fi
TARGET="$(cd "$TARGET" && pwd)"

# Guard against installing into the wrong folder: a real career-ops checkout has
# both the HTML builder and the base template.
if [ ! -f "$TARGET/build-cv-html.mjs" ] || [ ! -f "$TARGET/templates/cv-template.html" ]; then
	echo "$TARGET does not look like a career-ops checkout" >&2
	echo "(expected build-cv-html.mjs and templates/cv-template.html)" >&2
	exit 1
fi

# run <description> <command...> — executes, or just reports under --dry-run.
run() {
	local what="$1"
	shift
	if [ "$DRY_RUN" -eq 1 ]; then
		printf '  would %s\n' "$what"
	else
		"$@"
		printf '  %s\n' "$what"
	fi
}

# Same, but silent about the individual steps that make up one reported action.
quietly() {
	if [ "$DRY_RUN" -eq 0 ]; then
		"$@"
	fi
}

echo "Installing CV kit into $TARGET"
if [ "$DRY_RUN" -eq 1 ]; then
	echo "(dry run — nothing will be written)"
fi

# 1. Templates -------------------------------------------------------------
# Named variants only. The base cv-template.html is career-ops', never touched.
for tpl in "$KIT_DIR"/templates/cv-template.*.html; do
	name="$(basename "$tpl")"
	run "templates/$name" cp "$tpl" "$TARGET/templates/$name"
done

# 2. Word builder ----------------------------------------------------------
run "build-cv-docx.mjs" cp "$KIT_DIR/build-cv-docx.mjs" "$TARGET/build-cv-docx.mjs"

# 3. Skill -----------------------------------------------------------------
# career-ops keeps the real file in .agents/skills/ and symlinks it into each
# CLI's skills dir. Mirror that so the skill loads whichever CLI is in use.
quietly mkdir -p "$TARGET/.agents/skills/cv-word"
run ".agents/skills/cv-word/SKILL.md" \
	cp "$KIT_DIR/skills/cv-word/SKILL.md" "$TARGET/.agents/skills/cv-word/SKILL.md"

for cli in .claude .cursor .opencode .grok .kimi .qwen .antigravitycli; do
	# Only wire up CLIs this checkout already knows about.
	if [ ! -d "$TARGET/$cli/skills" ]; then
		continue
	fi
	quietly mkdir -p "$TARGET/$cli/skills/cv-word"
	run "$cli/skills/cv-word/SKILL.md -> .agents/skills/cv-word/SKILL.md" \
		ln -sf "../../../.agents/skills/cv-word/SKILL.md" "$TARGET/$cli/skills/cv-word/SKILL.md"
done

# 4. Dependency ------------------------------------------------------------
# Only the Word builder needs it; the HTML templates have no dependencies.
if [ "$NO_NPM" -eq 1 ]; then
	echo "  skipped npm install (--no-npm) — run: npm install --save-dev docx"
elif [ "$DRY_RUN" -eq 1 ]; then
	echo "  would run npm install --save-dev docx"
else
	echo "Installing the docx dependency..."
	if ! (cd "$TARGET" && npm install --save-dev docx --silent); then
		echo "npm install failed — run 'npm install --save-dev docx' in $TARGET yourself" >&2
	fi
fi

if [ "$DRY_RUN" -eq 1 ]; then
	echo "Dry run complete."
	exit 0
fi

cat <<EOF

Done. Check it worked:
  cd $TARGET
  node cv-templates.mjs list cv        # lists Compact, Executive, Jake, Leadership, Modern
  node build-cv-docx.mjs --self-test   # prints: build-cv-docx self-test passed

Then ask your AI CLI: "generate my CV with the modern template".
EOF
