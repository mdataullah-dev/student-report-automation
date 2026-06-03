/* Builds the IDP report template (.docx) — v2.
 * Clean (no highlights), list-token paragraphs the script turns into bullets,
 * embedded static 12-month timeline image, static guidelines + disclaimer.
 * Run: NODE_PATH=$(npm root -g) node generate_template.js
 * Output: IDP_Report_Template.docx
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, PageOrientation,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageBreak,
} = require('docx');

const NAVY = '1F3864';
const RED = 'C00000';
const CONTENT_W = 14400; // landscape US Letter, 0.5" margins

const logoNavy = fs.readFileSync('assets/logo_navy.png');
const logoWhite = fs.readFileSync('assets/logo.png');
const timeline = fs.readFileSync('assets/timeline.png');

function t(text, opts) { return new TextRun(Object.assign({ text: text }, opts || {})); }
function img(data, w, h, alt) {
  return new ImageRun({ type: 'png', data: data, transformation: { width: w, height: h },
    altText: { title: alt || 'image', description: alt || 'image', name: alt || 'image' } });
}
function heading(text) {
  return new Paragraph({ spacing: { before: 220, after: 120 },
    children: [t(text, { bold: true, size: 28, color: NAVY })] });
}
// A plain paragraph that holds a single list-token; the script replaces it with bullets.
function listToken(token) { return new Paragraph({ children: [t(token)] }); }
// Static bullet (kept as-is in the report).
function bullet(text) {
  return new Paragraph({ bullet: { level: 0 }, spacing: { after: 40 }, children: [t(text)] });
}
function cell(children, opts) {
  opts = opts || {};
  return new TableCell({
    width: { size: opts.w, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    borders: cellBorders,
    children: children,
  });
}
const cb = { style: BorderStyle.SINGLE, size: 1, color: '9CA3AF' };
const cellBorders = { top: cb, bottom: cb, left: cb, right: cb };

// running header / footer
const header = new Header({ children: [ new Paragraph({
  alignment: AlignmentType.RIGHT, spacing: { after: 0 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 4 } },
  children: [ img(logoNavy, 74, 70, 'ECS logo') ],
}) ] });
const footer = new Footer({ children: [ new Paragraph({
  alignment: AlignmentType.RIGHT,
  children: [ t('Confidential', { color: RED, size: 16, italics: true }) ],
}) ] });

// ---------- PAGE 1: cover ----------
const coverInner = [];
for (let i = 0; i < 3; i++) coverInner.push(new Paragraph({ children: [t('')] }));
coverInner.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [img(logoWhite, 260, 246, 'ECS logo')] }));
for (let i = 0; i < 4; i++) coverInner.push(new Paragraph({ children: [t('')] }));
coverInner.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [ t('Date: ', { bold: true, color: 'FFFFFF', size: 24 }), t('{{report_date}}', { color: 'FFFFFF', size: 24 }) ] }));
coverInner.push(new Paragraph({ alignment: AlignmentType.CENTER,
  children: [ t('Name: ', { bold: true, color: 'FFFFFF', size: 24 }), t('{{student_name}}', { color: 'FFFFFF', size: 24 }) ] }));
for (let i = 0; i < 3; i++) coverInner.push(new Paragraph({ children: [t('')] }));

const coverTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [ new TableRow({ children: [
    new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: NAVY },
      margins: { top: 200, bottom: 200, left: 200, right: 200 }, children: coverInner }) ] }) ],
});

// ---------- PAGE 2: profile + programs table + skills ----------
const profileRow = new Paragraph({ spacing: { after: 160 }, children: [
  t('Department: ', { bold: true }), t('{{department}}'), t('     '),
  t('Job title: ', { bold: true }), t('{{job_title}}'), t('     '),
  t('Experience: ', { bold: true }), t('{{experience}}'), t('     '),
  t('Preferred language: ', { bold: true }), t('{{preferred_language}}'), t('     '),
  t('English level: ', { bold: true }), t('{{english_level}}'),
] });

const tblCols = [800, 5600, 1800, 2600, 3600];
const headerCells = ['#', 'Course / Certification', 'Duration', 'Type', 'Action'].map(function (label, i) {
  return cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [t(label, { bold: true, color: 'FFFFFF' })] })],
    { w: tblCols[i], fill: NAVY });
});
const bodyCells = [
  cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [t('1')] })], { w: tblCols[0] }),
  cell([new Paragraph({ children: [t('{{prog_name}}')] })], { w: tblCols[1] }),
  cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [t('{{prog_duration}}')] })], { w: tblCols[2] }),
  cell([new Paragraph({ children: [t('{{prog_type}}')] })], { w: tblCols[3] }),
  cell([new Paragraph({ children: [t('{{prog_action}}')] })], { w: tblCols[4] }),
];
const programsTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: tblCols,
  rows: [ new TableRow({ tableHeader: true, children: headerCells }), new TableRow({ children: bodyCells }) ],
});

const page2 = [
  profileRow,
  heading('Proposed development plan:'),
  programsTable,
  new Paragraph({ children: [t('')] }),
  heading('The skills that will be acquired:'),
  listToken('{{skills_list}}'),
];

// ---------- PAGE 3: timeline ----------
const page3 = [
  heading('Proposed timeline'),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 },
    children: [ img(timeline, 940, 188, '12-month timeline') ] }),
  new Paragraph({ spacing: { before: 80, after: 40 }, children: [ t('Phase 1: ', { bold: true, color: NAVY }), t('{{phase1_label}}', { bold: true, color: NAVY }) ] }),
  listToken('{{phase1_list}}'),
  new Paragraph({ spacing: { before: 80, after: 40 }, children: [ t('Phase 2: ', { bold: true, color: NAVY }), t('{{phase2_label}}', { bold: true, color: NAVY }) ] }),
  listToken('{{phase2_list}}'),
  new Paragraph({ spacing: { before: 80, after: 40 }, children: [ t('Phase 3: ', { bold: true, color: NAVY }), t('{{phase3_label}}', { bold: true, color: NAVY }) ] }),
  listToken('{{phase3_list}}'),
];

// ---------- PAGE 4: outlook + guidelines + disclaimer ----------
const page4 = [
  heading('Improve the route (optional but important):'),
  listToken('{{improve_list}}'),
  heading('After completing this path, you will be able to:'),
  listToken('{{outcome_list}}'),
  heading('General Guidelines'),
  bullet('Commit to the plan while applying learning directly at work.'),
  bullet('Focus on building practical skills, not only theoretical knowledge.'),
  bullet('Strengthen communication with suppliers and stakeholders.'),
  bullet('Track progress through measurable KPIs.'),
  bullet('Stay updated with industry trends and best practices.'),
  heading('Competitiveness Enhancement'),
  bullet('Combine your background with new domain expertise.'),
  bullet('Develop strong negotiation and stakeholder-management skills.'),
  bullet('Use digital tools and AI to support decision-making.'),
  bullet('Improve cost optimization and value creation.'),
  bullet('Build a strong professional profile in your field.'),
  heading('Disclaimer'),
  new Paragraph({ children: [ t(
    'This plan is based on professional development best practices and training programs provided ' +
    'by the center. Participation in these programs is not mandatory for competency development, and ' +
    'outcomes depend on individual performance, experience, and market conditions.') ] }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
  sections: [{
    properties: { page: {
      size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE },
      margin: { top: 1080, right: 720, bottom: 720, left: 720 },
    } },
    headers: { default: header },
    footers: { default: footer },
    children: [].concat(
      [coverTable, new Paragraph({ children: [new PageBreak()] })],
      page2, [new Paragraph({ children: [new PageBreak()] })],
      page3, [new Paragraph({ children: [new PageBreak()] })],
      page4,
    ),
  }],
});

Packer.toBuffer(doc).then(function (buf) {
  fs.writeFileSync('IDP_Report_Template.docx', buf);
  console.log('wrote IDP_Report_Template.docx', buf.length, 'bytes');
});
