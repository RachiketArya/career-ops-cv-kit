---
name: cv-word
description: >-
  Produce an editable Word (.docx) CV with hand-tuned, measured formatting. Use
  when the user asks for a Word/Doc/docx CV, says an employer or ATS portal only
  accepts .doc/.docx, wants to hand-edit a generated CV, or asks for "the Word
  version" of a CV that was generated as a PDF.
user_invocable: true
user-invocable: true
argument-hint: "[rich-payload.json]"
license: MIT
---

# CV → Word (.docx)

Builds the Word CV with the [`docx`](https://docx.js.org) npm library, following
the official Anthropic docx skill
(`github.com/anthropics/skills/tree/main/skills/docx`).

Installed by [career-ops-cv-kit](https://github.com/RachiketArya/career-ops-cv-kit),
not shipped with career-ops. After `node update-system.mjs apply`, re-run the
kit's `install.sh` if anything here goes missing.

## The commands

```bash
# Build the .docx AND re-derive the flat payload the HTML/PDF templates use
node build-cv-docx.mjs output/cv-payload-rich.json output/CV.docx \
  --emit-flat output/cv-payload-master.json

node build-cv-docx.mjs --self-test
```

Name the output after the candidate (`output/Jane_Smith_CV.docx`) so several
tailored versions can coexist.

Then validate visually — this step is not optional:

```bash
soffice --headless --convert-to pdf --outdir /tmp output/CV.docx
pdftoppm -png -r 110 /tmp/CV.pdf /tmp/page   # then LOOK at the pages
```

`soffice` renders with Carlito, the metric-compatible Calibri clone, so
pagination is representative even though the glyphs aren't identical. Word uses
real Calibri.

## Two payloads, one source

- **`output/cv-payload-rich.json` is the source of truth for CV rendering.** It
  models what the flat payload cannot: a company block with a descriptor line
  and multiple sub-roles that each own their bullets, plus inline `**bold**`
  anywhere in a bullet.
- **`--emit-flat` regenerates the flat payload** that `build-cv-html.mjs` and the
  PDF templates consume — stripping `**` markers and splitting sub-roles into
  separate entries. **Always pass `--emit-flat`** so the Word and PDF versions
  cannot drift. Never hand-edit the flat payload; edit the rich one and re-emit.

Content in either payload must trace back to `cv.md` / `config/profile.yml` /
`article-digest.md` per career-ops' Source-of-Truth Boundary. Generating a Word
file is not a licence to add claims that aren't in those files.

## Verify facts before sending

`verify-cv-facts.mjs` reads HTML, not docx, so check the HTML twin built from the
emitted flat payload:

```bash
node build-cv-html.mjs output/cv-payload-master.json output/_check.html
node verify-cv-facts.mjs output/_check.html
```

It matches on near-exact phrasing, so a reworded bullet fails even when the claim
is identical. Fix by matching `cv.md`'s wording, not by adding an exception.

## Formatting spec — measured, do not guess

Taken off a hand-tuned Google-Docs CV with `pdffonts` and `pdftotext -bbox`. The
generated file reproduces these; re-measure before changing any of them.

| Element | Spec |
|---|---|
| Font | Calibri (original embeds Calibri + Bold/Italic/BoldItalic) |
| Page | A4, margins 39.6pt = 0.55in all round (792 twips) |
| Name | 19pt bold, `#1F3864`, uppercase |
| Contact | 9.5pt, `#595959`, ` • ` separators |
| Section heading | 11pt bold, `#1F3864`, uppercase, bottom rule |
| Company header | 11pt bold black, then ` \| ` + dates 10pt bold `#1F3864`, **inline** — the original does *not* right-align dates with a tab stop |
| Descriptor line | 9.5pt italic `#595959` |
| Role sub-heading | 10pt bold italic `#1F3864` |
| Body / bullets | 10pt, **justified**, single spacing (12.2pt line box) |
| Bullet indent | left 330 twips, hanging 210 (glyph lands 120 in) |

Accent `#1F3864` and grey `#595959` are the only two colours besides black. To
rebrand, change `ACCENT` and `GREY` at the top of `build-cv-docx.mjs`.

## Gotchas that cost time

- **Bullets must use the formal `numbering` config** with `LevelFormat.BULLET` —
  never a literal `•` character in the text.
- **Justification matters.** The original justifies bullet text; left-aligned
  output looks visibly different (ragged right vs stretched inter-word spacing).
- **Measure page fit at the real content-box width.** A default browser viewport
  wraps fewer lines and understates height.
- **No inline markdown links in bullets.** Only `**bold**` is parsed; hyperlinks
  work solely in the contact block via `candidate.links`. A `[label](url)` inside
  a bullet renders literally — write URLs as bare text.
- `docx` is a **devDependency**. `update-system.mjs` may overwrite `package.json`
  on a career-ops update — if the build then fails with a missing module, re-run
  `npm install --save-dev docx`.
