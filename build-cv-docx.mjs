#!/usr/bin/env node

// build-cv-docx.mjs — Word (.docx) builder for career-ops CVs.
//
// Follows the official Anthropic `docx` skill
// (github.com/anthropics/skills/tree/main/skills/docx): use the `docx` npm
// library, formal numbering config for bullets (never literal "•" characters),
// separate Paragraphs instead of "\n", then convert to PDF and rasterise the
// pages to check pagination visually before shipping.
//
// ── Formatting spec ────────────────────────────────────────────────────────
// Measured directly off Rachiket's own Google-Docs CV
// (Downloads/Rachiket_Arya_CV.pdf) with `pdffonts` and `pdftotext -bbox`, so the
// generated file matches the document he hand-tuned rather than approximating
// it. Do not "tidy" these numbers without re-measuring:
//
//   font            Calibri (original embeds Calibri + Calibri-Bold/Italic/BoldItalic)
//   page            A4 596x842pt; margins 39.6pt = 0.55in on all sides
//   name            19pt bold, #1F3864, uppercase
//   contact         9.5pt, #595959, " • " separators
//   section head    11pt bold, #1F3864, uppercase, bottom rule
//   company head    11pt bold black, then " | " + dates 10pt bold #1F3864 INLINE
//                   (the original does not right-align dates with a tab stop)
//   descriptor      9.5pt italic #595959
//   role sub-head   10pt bold italic #1F3864
//   body / bullets  10pt, single spacing (12.2pt measured for 10pt Calibri)
//   bullet indent   left 330 twips, hanging 210 (glyph lands 120 twips in)
//
// ── Payload ────────────────────────────────────────────────────────────────
// This builder consumes the RICH payload (output/cv-payload-rich.json), which
// models what the flat build-cv-html.mjs payload cannot: a company block with a
// descriptor line and multiple sub-roles each owning their own bullets (Aspire →
// Head of AI + Senior AI Product Engineer; Jackett / CoTeach → CoTeach +
// Jackett), plus inline **bold** anywhere in a bullet.
//
// To stop the Word and PDF versions drifting, `--emit-flat <path>` derives the
// flat HTML payload from this same rich source: it strips the ** markers and
// splits sub-roles into separate entries. Generate both from one file.
//
// Usage:
//   node build-cv-docx.mjs <rich-payload.json> <out.docx> [--emit-flat <flat.json>]
//   node build-cv-docx.mjs --self-test

import { readFile, writeFile } from 'fs/promises';
import { resolve, dirname, isAbsolute } from 'path';
import { fileURLToPath } from 'url';
import {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  AlignmentType, BorderStyle, LevelFormat, HeadingLevel,
} from 'docx';

const __dirname = dirname(fileURLToPath(import.meta.url));

const FONT = 'Calibri';
const ACCENT = '1F3864';
const GREY = '595959';
const BLACK = '000000';

// half-points (docx `size` unit): 10pt -> 20
const SZ = { name: 38, contact: 19, section: 22, company: 22, dates: 20, descriptor: 19, role: 20, body: 20 };

const MARGIN = 792;            // 0.55in in twips, matching the original's 39.6pt
const PAGE_W = 11906;          // A4
const PAGE_H = 16838;
const BULLET_INDENT = { left: 330, hanging: 210 };

// ── inline **bold** → runs ──────────────────────────────────────────────────
// The payload carries markdown-style emphasis because the bold lead-ins are part
// of the CV's voice ("Built the AI Centre of Excellence..." then prose). Word
// needs these as separate runs, so the marker is parsed rather than escaped.

export function parseBold(text, base = {}) {
  const out = [];
  const re = /\*\*(.+?)\*\*/gs;
  let last = 0, m;
  while ((m = re.exec(String(text ?? ''))) !== null) {
    if (m.index > last) out.push({ text: String(text).slice(last, m.index), bold: false });
    out.push({ text: m[1], bold: true });
    last = m.index + m[0].length;
  }
  if (last < String(text ?? '').length) out.push({ text: String(text).slice(last), bold: false });
  return out
    .filter((r) => r.text !== '')
    .map((r) => new TextRun({ ...base, text: r.text, bold: r.bold || base.bold === true }));
}

export function stripBold(text) {
  return String(text ?? '').replace(/\*\*(.+?)\*\*/gs, '$1');
}

// ── paragraph builders ──────────────────────────────────────────────────────

function nameP(name) {
  return new Paragraph({
    spacing: { after: 40, line: 240, lineRule: 'auto' },
    children: [new TextRun({ text: name, bold: true, size: SZ.name, color: ACCENT, font: FONT })],
  });
}

function contactP(candidate) {
  const items = candidate.contact || [];
  const links = candidate.links || {};
  const children = [];
  items.forEach((item, i) => {
    if (i > 0) {
      children.push(new TextRun({ text: '  •  ', size: SZ.contact, color: GREY, font: FONT }));
    }
    const style = { size: SZ.contact, color: GREY, font: FONT };
    if (links[item]) {
      children.push(new ExternalHyperlink({
        link: links[item],
        children: [new TextRun({ ...style, text: item })],
      }));
    } else {
      children.push(new TextRun({ ...style, text: item }));
    }
  });
  return new Paragraph({ spacing: { after: 200, line: 240, lineRule: 'auto' }, children });
}

function sectionHeadP(title) {
  return new Paragraph({
    // HeadingLevel keeps Word's navigation pane and any TOC working; the visual
    // styling is carried by the run + border, per the skill's note that custom
    // looks still need a built-in heading to be outline-visible.
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 200, after: 120, line: 240, lineRule: 'auto' },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } },
    children: [new TextRun({ text: title, bold: true, size: SZ.section, color: ACCENT, font: FONT })],
  });
}

function companyHeadP(header, dates) {
  const children = [new TextRun({ text: header, bold: true, size: SZ.company, color: BLACK, font: FONT })];
  if (dates) {
    children.push(new TextRun({ text: ' | ', bold: true, size: SZ.company, color: BLACK, font: FONT }));
    children.push(new TextRun({ text: dates, bold: true, size: SZ.dates, color: ACCENT, font: FONT }));
  }
  return new Paragraph({
    spacing: { before: 180, after: 0, line: 240, lineRule: 'auto' },
    keepNext: true,
    children,
  });
}

function descriptorP(text) {
  return new Paragraph({
    spacing: { before: 20, after: 60, line: 240, lineRule: 'auto' },
    keepNext: true,
    children: [new TextRun({ text, italics: true, size: SZ.descriptor, color: GREY, font: FONT })],
  });
}

function roleP(title) {
  return new Paragraph({
    spacing: { before: 100, after: 40, line: 240, lineRule: 'auto' },
    keepNext: true,
    children: [new TextRun({ text: title, bold: true, italics: true, size: SZ.role, color: ACCENT, font: FONT })],
  });
}

function bulletP(text) {
  return new Paragraph({
    numbering: { reference: 'cv-bullets', level: 0 },
    // The original justifies bullet text — visible as stretched inter-word
    // spacing on wrapped lines ("Drove  the  secure  deployment  of  key…").
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 80, line: 240, lineRule: 'auto' },
    indent: BULLET_INDENT,
    children: parseBold(text, { size: SZ.body, color: BLACK, font: FONT }),
  });
}

function proseP(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 80, line: 240, lineRule: 'auto' },
    children: parseBold(text, { size: SZ.body, color: BLACK, font: FONT }),
  });
}

// ── document assembly ───────────────────────────────────────────────────────

export function buildChildren(payload) {
  const c = payload.candidate || {};
  const out = [nameP(c.name || ''), contactP(c)];

  for (const sec of payload.sections || []) {
    if (!sec) continue;
    out.push(sectionHeadP(sec.title || ''));
    if (sec.prose) out.push(proseP(sec.prose));
    if (sec.descriptor) out.push(descriptorP(sec.descriptor));
    for (const b of sec.bullets || []) if (b) out.push(bulletP(b));

    for (const e of sec.entries || []) {
      if (!e) continue;
      out.push(companyHeadP(e.header || '', e.dates));
      if (e.descriptor) out.push(descriptorP(e.descriptor));
      for (const b of e.bullets || []) if (b) out.push(bulletP(b));
      for (const r of e.roles || []) {
        if (!r) continue;
        out.push(roleP(r.title || ''));
        for (const b of r.bullets || []) if (b) out.push(bulletP(b));
      }
    }
  }
  return out;
}

export function buildDoc(payload) {
  return new Document({
    creator: stripBold(payload?.candidate?.name || 'CV'),
    title: `${stripBold(payload?.candidate?.name || '')} - CV`,
    styles: {
      default: {
        document: { run: { font: FONT, size: SZ.body }, paragraph: { spacing: { line: 240, lineRule: 'auto' } } },
        // Neutralise Word's built-in Heading 1 look; the run carries the styling.
        heading1: { run: { font: FONT, size: SZ.section, bold: true, color: ACCENT } },
      },
      characterStyles: [
        // Word underlines hyperlinks blue by default, which fights the grey
        // contact line in the original.
        { id: 'Hyperlink', name: 'Hyperlink', run: { color: GREY, underline: undefined } },
      ],
    },
    numbering: {
      config: [{
        reference: 'cv-bullets',
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: BULLET_INDENT }, run: { font: FONT, size: SZ.body } },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      children: buildChildren(payload),
    }],
  });
}

// ── rich → flat (keeps the HTML/PDF pipeline in sync) ───────────────────────

export function flattenPayload(rich) {
  const c = rich.candidate || {};
  const contact = c.contact || [];
  const links = c.links || {};
  const emailItem = contact.find((x) => /@/.test(x));
  const liItem = contact.find((x) => /linkedin/i.test(x));
  const phoneItem = contact.find((x) => /^\+?[\d\s()-]{7,}$/.test(x));
  const locItem = contact.find((x) => x !== emailItem && x !== liItem && x !== phoneItem);

  const flat = {
    lang: 'en',
    page_format: 'a4',
    sections: {},
    candidate: {
      // Title-case the name: the rich payload uppercases it for the Word header,
      // but the HTML templates apply their own text-transform.
      name: stripBold(c.name || '').replace(/\b([A-Z])([A-Z]+)\b/g, (_, a, b) => a + b.toLowerCase()),
      phone: phoneItem || '',
      email: emailItem || '',
      location: locItem || '',
    },
    summary: '',
    competencies: [],
    experience: [],
    projects: [],
    education: [],
    skills: [],
  };
  if (liItem) flat.candidate.linkedin = { url: links[liItem] || `https://${liItem}`, display: liItem };

  for (const sec of rich.sections || []) {
    const title = (sec.title || '').toUpperCase();
    if (sec.prose && /SUMMARY/.test(title)) { flat.summary = stripBold(sec.prose); continue; }

    if (/EXPERIENCE/.test(title)) {
      for (const e of sec.entries || []) {
        const { header, location } = splitHeader(e.header || '');
        const common = { company: header, location, dates: e.dates ? e.dates.replace(/–/g, '-') : '' };
        const lead = e.descriptor ? [stripBold(e.descriptor)] : [];
        if (e.roles?.length) {
          for (const r of e.roles) {
            flat.experience.push({
              ...common,
              role: r.title || '',
              bullets: [...(flat.experience.length === 0 ? lead : []), ...(r.bullets || []).map(stripBold)],
            });
          }
        } else {
          flat.experience.push({
            ...common,
            role: roleFromHeader(e.header || ''),
            bullets: [...lead, ...(e.bullets || []).map(stripBold)],
          });
        }
      }
      continue;
    }

    if (/EDUCATION/.test(title)) {
      for (const e of sec.entries || []) {
        const [t, org] = (e.header || '').split(/\s+-\s+/);
        flat.education.push({
          title: t || e.header || '',
          org: org || '',
          year: '',
          description: (e.bullets || []).map(stripBold).join(' '),
        });
      }
      continue;
    }

    if (/SPEAKING|FACILITATION/.test(title)) {
      flat.sections.projects = titleCaseSection(sec.title);
      for (const b of sec.bullets || []) {
        const m = /^\*\*(.+?)\*\*\s*-?\s*(.*)$/s.exec(b);
        flat.projects.push(m ? { name: m[1], description: m[2] } : { name: stripBold(b) });
      }
      continue;
    }

    if (/ADDITIONAL/.test(title)) {
      flat.sections.skills = titleCaseSection(sec.title);
      for (const b of sec.bullets || []) {
        const m = /^\*\*(.+?):\*\*\s*(.*)$/s.exec(b);
        if (m) flat.skills.push({ category: m[1], items: [stripBold(m[2])] });
        else flat.skills.push({ category: '', items: [stripBold(b)] });
      }
      continue;
    }
  }

  // The HTML templates render a Core Competencies strip unconditionally, and
  // cv-sections-core.mjs only strips projects/education/certifications — so an
  // empty list would leave a bare heading. Derive it from the Skills line.
  const skillsLine = flat.skills.find((s) => /^skills$/i.test(s.category));
  if (skillsLine) {
    flat.competencies = String(skillsLine.items[0] || '')
      // Strip parentheticals FIRST — splitting on commas while "(Claude, Claude
      // Code, LangGraph)" is still present shatters one skill into fragments
      // like "LLM & agent orchestration (Claude" and "Agent-harness)".
      .replace(/\s*\([^)]*\)/g, '')
      .split(/,\s*/)
      // Preserve original casing: capitalising the first letter turns "n8n at
      // enterprise scale" into "N8n at enterprise scale".
      .map((s) => s.replace(/\.$/, '').trim())
      .filter((s) => s && s.length < 34)
      .slice(0, 10);
  }
  return flat;
}

function splitHeader(header) {
  const m = /^(.*?)\s*\(([^)]+)\)/.exec(header);
  if (m) return { header: m[1].trim(), location: m[2].trim() };
  return { header: header.replace(/\s*-\s*.*$/, '').trim(), location: '' };
}

function roleFromHeader(header) {
  const m = /-\s*(.+)$/.exec(header.replace(/\([^)]*\)/g, ''));
  return m ? m[1].trim() : '';
}

function titleCaseSection(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\b\w/g, (ch) => ch.toUpperCase())
    .replace(/\bAnd\b/g, 'and');
}

// ── CLI ─────────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.includes('--self-test')) return selfTest();

  const flatIdx = args.indexOf('--emit-flat');
  const flatOut = flatIdx >= 0 ? args[flatIdx + 1] : null;
  const positional = args.filter((a, i) => !a.startsWith('--') && i !== flatIdx + 1);

  if (positional.length < 2) {
    console.error('Usage: node build-cv-docx.mjs <rich-payload.json> <out.docx> [--emit-flat <flat.json>]');
    console.error('       node build-cv-docx.mjs --self-test');
    process.exitCode = 1;
    return;
  }

  const abs = (p) => (isAbsolute(p) ? p : resolve(process.cwd(), p));
  const outPath = abs(positional[1]);
  // Same containment rule as generate-pdf.mjs: never write outside the project.
  if (!outPath.startsWith(resolve(__dirname))) {
    console.error(`Refusing to write outside the project directory: ${outPath}`);
    process.exitCode = 1;
    return;
  }

  const payload = JSON.parse(await readFile(abs(positional[0]), 'utf8'));
  const buf = await Packer.toBuffer(buildDoc(payload));
  await writeFile(outPath, buf);

  const result = { file: outPath, sizeKB: +(buf.length / 1024).toFixed(1) };
  if (flatOut) {
    const flatPath = abs(flatOut);
    if (!flatPath.startsWith(resolve(__dirname))) {
      console.error(`Refusing to write outside the project directory: ${flatPath}`);
      process.exitCode = 1;
      return;
    }
    const flat = flattenPayload(payload);
    await writeFile(flatPath, JSON.stringify(flat, null, 2) + '\n');
    result.flat = flatPath;
    result.flatCounts = {
      experienceEntries: flat.experience.length,
      totalBullets: flat.experience.reduce((n, e) => n + e.bullets.length, 0),
      projectEntries: flat.projects.length,
      skillCategories: flat.skills.length,
      competencies: flat.competencies.length,
    };
  }
  console.log(JSON.stringify(result, null, 2));
}

function selfTest() {
  const fails = [];
  const runs = parseBold('**Lead in.** then prose', { size: 20 });
  if (runs.length !== 2) fails.push(`parseBold split into ${runs.length} runs, expected 2`);
  if (stripBold('a **b** c') !== 'a b c') fails.push('stripBold left markers behind');
  if (parseBold('no markers').length !== 1) fails.push('unmarked text should be a single run');
  if (parseBold('**a** mid **b**').length !== 3) fails.push('multiple bold spans not handled');

  const rich = {
    candidate: { name: 'TEST USER', contact: ['Singapore', 't@e.com', 'linkedin.com/in/x'], links: { 't@e.com': 'mailto:t@e.com' } },
    sections: [
      { title: 'SUMMARY', prose: 'Summary text.' },
      { title: 'WORK EXPERIENCE', entries: [
        { header: 'Aspire (Singapore)', dates: 'Jun 2025 – Present', descriptor: 'Desc.',
          roles: [{ title: 'Head of AI', bullets: ['**Bold.** rest'] }, { title: 'Engineer', bullets: ['Second role bullet'] }] },
      ] },
      { title: 'ADDITIONAL INFORMATION', bullets: ['**Skills:** Python, SQL, MCP server design'] },
    ],
  };
  const children = buildChildren(rich);
  if (children.length < 8) fails.push(`buildChildren produced only ${children.length} paragraphs`);

  const flat = flattenPayload(rich);
  if (flat.experience.length !== 2) fails.push(`flatten should split 2 sub-roles, got ${flat.experience.length}`);
  if (flat.experience[0].company !== 'Aspire') fails.push(`company parsed as "${flat.experience[0].company}"`);
  if (flat.experience[0].location !== 'Singapore') fails.push('location not split from header');
  if (flat.experience[0].dates !== 'Jun 2025 - Present') fails.push('en dash not normalised in dates');
  if (flat.experience[0].bullets.some((b) => b.includes('**'))) fails.push('flatten leaked ** markers');
  if (flat.candidate.name !== 'Test User') fails.push(`name not de-uppercased: "${flat.candidate.name}"`);
  if (!flat.competencies.length) fails.push('competencies not derived (HTML template would show a bare heading)');
  if (!flat.summary) fails.push('summary not carried over');

  if (fails.length) {
    for (const f of fails) console.error(`Self-test failed: ${f}`);
    process.exitCode = 1;
  } else {
    console.log('build-cv-docx self-test passed');
  }
}

main().catch((err) => {
  console.error(err?.stack || err?.message || err);
  process.exitCode = 1;
});
