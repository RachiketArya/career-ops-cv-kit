# career-ops CV kit

Five extra CV templates and a Word (`.docx`) builder for
[career-ops](https://github.com/santifer/career-ops).

It installs **alongside** career-ops rather than replacing anything in it, so
you keep pulling upstream updates normally. Nothing here contains personal data
— the templates are placeholder-driven and the builder reads whatever payload
your career-ops instance generates from your own `cv.md`.

## What you get

| Template | Looks like | Good for |
|---|---|---|
| `modern` | Oversized name, accent bars, tinted summary panel | Product and tech roles |
| `compact` | Tight leading, left-rail headings | Squeezing two pages into one |
| `executive` | Serif, centred header, no colour | Banks, funds, traditional enterprises |
| `leadership` | Competencies block ahead of the chronology | Senior/leadership applications |
| `jake` | HTML port of "Jake's Resume" | The classic engineering resume look |

All five keep career-ops' ATS discipline: single column, static system font
stack, ligatures disabled. Decorative CSS never reorders the text flow, so PDF
text extraction stays clean.

Plus `build-cv-docx.mjs`, which produces an **editable Word file** — for the
employers and ATS portals that reject PDFs — and the `cv-word` skill that tells
your AI CLI how to drive it.

## Install

You need a working career-ops checkout first (its own README covers that).

```bash
git clone https://github.com/RachiketArya/career-ops-cv-kit.git
cd career-ops-cv-kit
./install.sh /path/to/career-ops
```

Pass `--dry-run` to see what it would do, `--no-npm` to skip installing the
`docx` dependency. With no path, it looks for a `career-ops` folder next to this
one.

## Use

Pick a template per CV, in conversation with your AI CLI:

> generate my CV with the **executive** template

Or make one the default for every CV, in `config/profile.yml`:

```yaml
cv:
  template: modern
```

Under the hood these are career-ops' own commands — the kit adds files, not a
new pipeline:

```bash
node cv-templates.mjs list cv          # what's available
node cv-templates.mjs resolve cv modern # path to fill
```

For the Word version, ask for "a Word/docx CV" and the `cv-word` skill takes
over. Directly:

```bash
node build-cv-docx.mjs output/cv-payload-rich.json output/Jane_Smith_CV.docx \
  --emit-flat output/cv-payload-master.json
```

`--emit-flat` re-derives the payload the HTML/PDF templates use from the same
source, so your Word and PDF CVs can never drift apart. Always pass it.

## Keeping everything up to date

Two independent update paths, which is the point of shipping this separately:

```bash
# career-ops itself — unaffected by the kit
cd /path/to/career-ops && node update-system.mjs check   # then: apply

# this kit
cd /path/to/career-ops-cv-kit && git pull && ./install.sh /path/to/career-ops
```

A career-ops update rewrites its own system files. It leaves the kit's
templates and the skill alone (they are not files it knows about), but it can
reset `package.json` and drop the `docx` devDependency. If a Word build then
fails with a missing module:

```bash
npm install --save-dev docx
```

Re-running `./install.sh` is always safe and fixes anything an update disturbed.

## Uninstall

```bash
cd /path/to/career-ops
rm templates/cv-template.{modern,compact,executive,leadership,jake}.html
rm build-cv-docx.mjs
rm -rf .agents/skills/cv-word */skills/cv-word
```

## Licence

MIT, same as career-ops. The `jake` template follows the design of
[Jake's Resume](https://github.com/jakegut/resume) (MIT) — this is an
independent HTML/CSS implementation, not a copy of its LaTeX source.
