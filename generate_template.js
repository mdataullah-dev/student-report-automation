/* Builds the starter IDP report template (.docx) with all {{tokens}} in place.
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
const CYAN = 'CFF5FF';   // dynamic fields (auto-filled by the script)
const GREEN = 'D9F2D0';  // staff-editable sections
const RED = 'C00000';
const CONTENT_W = 14400; // landscape US Letter, 0.5" margins

const logoNavy = fs.readFileSync('assets/logo_navy.png'); // white logo on navy (for white pages)
const logoWhite = fs.readFileSync('assets/logo.png');     // white logo (for navy cover)

// ---------- helpers ----------
function t(text, opts) { return new TextRun(Object.assign({ text: text }, opts || {})); }
function dyn(text) { return new TextRun({ text: text, shading: { type: ShadingType.CLEAR, fill: CYAN } }); }
function img(data, w, h) {
  return new ImageRun({ type: 'png', data: data, transformation: { width: w, height: h },
    altText: { title: 'ECS logo', description: 'ECS Elite Training Center', name: 'logo' } });
}
function heading(text) {
  return new Paragraph({ spacing: { before: 200, after: 120 },
    children: [t(text, { bold: true, size: 28, color: NAVY })] });
}
function bullet(runs, fill) {
  return new Paragraph({ bullet: { level: 0 }, spacing: { after: 40 },
    children: Array.isArray(runs) ? runs : [runs],
    shading: fill ? { type: ShadingType.CLEAR, fill: fill } : undefined });
}
function cell(children, opts) {
  opts = opts || {};
  return new TableCell({
    width: { size: opts.w, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: children,
  });
}
const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: '9CA3AF' };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

// ---------- running header / footer (pages 2-4) ----------
const header = new Header({ children: [ new Paragraph({
  alignment: AlignmentType.RIGHT, spacing: { after: 0 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 4 } },
  children: [ img(logoNavy, 74, 70) ],
}) ] });
const footer = new Footer({ children: [ new Paragraph({
  alignment: AlignmentType.RIGHT,
  children: [ t('Confidential', { color: RED, size: 16, italics: true }) ],
}) ] });

// ---------- PAGE 1: cover ----------
const coverInner = [];
for (let i = 0; i < 3; i++) coverInner.push(new Paragraph({ children: [t('')] }));
coverInner.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [img(logoWhite, 260, 246)] }));
for (let i = 0; i < 4; i++) coverInner.push(new Paragraph({ children: [t('')] }));
coverInner.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [ t('Date: ', { bold: true, color: 'FFFFFF', size: 24 }), dyn('{{report_date}}') ] }));
coverInner.push(new Paragraph({ alignment: AlignmentType.CENTER,
  children: [ t('Name: ', { bold: true, color: 'FFFFFF', size: 24 }), dyn('{{student_name}}') ] }));
for (let i = 0; i < 3; i++) coverInner.push(new Paragraph({ children: [t('')] }));

const coverTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: [ new TableRow({ children: [
    new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: NAVY },
      margins: { top: 200, bottom: 200, left: 200, right: 200 },
      children: coverInner }) ] }) ],
});

// ---------- PAGE 2: profile + programs table ----------
const profileRow = new Paragraph({ spacing: { after: 160 }, children: [
  t('Department: ', { bold: true }), dyn('{{department}}'), t('     '),
  t('Job title: ', { bold: true }), dyn('{{job_title}}'), t('     '),
  t('Experience: ', { bold: true }), dyn('{{experience}}'), t('     '),
  t('Preferred language: ', { bold: true }), dyn('{{preferred_language}}'), t('     '),
  t('English level: ', { bold: true }), dyn('{{english_level}}'),
] });

const tblCols = [800, 5600, 1800, 2600, 3600];
const headerCells = ['#', 'Course / Certification', 'Duration', 'Type', 'Action'].map(function (label, i) {
  return cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [t(label, { bold: true, color: 'FFFFFF' })] })],
    { w: tblCols[i], fill: NAVY });
});
const bodyCells = [
  cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [t('1')] })], { w: tblCols[0] }),
  cell([new Paragraph({ children: [dyn('{{prog_name}}')] })], { w: tblCols[1] }),
  cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [dyn('{{prog_duration}}')] })], { w: tblCols[2] }),
  cell([new Paragraph({ children: [dyn('{{prog_type}}')] })], { w: tblCols[3] }),
  cell([new Paragraph({ children: [dyn('{{prog_action}}')] })], { w: tblCols[4] }),
].map(function (c) { c.options.borders = cellBorders; return c; });
const headerCellsB = headerCells.map(function (c) { c.options.borders = cellBorders; return c; });

const programsTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: tblCols,
  rows: [ new TableRow({ tableHeader: true, children: headerCellsB }), new TableRow({ children: bodyCells }) ],
});

const page2 = [
  profileRow,
  heading('Proposed development plan:'),
  programsTable,
  new Paragraph({ children: [t('')] }),
  heading('The skills that will be acquired:'),
  bullet([t('Staff: list the key skills the student will gain (editable).')], GREEN),
  bullet([t('Add one bullet per outcome as needed.')], GREEN),
];

// ---------- PAGE 3: timeline ----------
function phaseBlock(labelTok, compTok) {
  return [
    new Paragraph({ spacing: { before: 160, after: 60 }, children: [
      t('Phase: ', { bold: true, color: NAVY }), dyn(labelTok) ] }),
    new Paragraph({ children: [ dyn(compTok) ] }),
  ];
}
const page3 = [
  heading('Proposed timeline'),
  new Paragraph({ spacing: { after: 120 }, children: [
    t('[Keep the 12-month timeline graphic here as a static image/drawing — staff design once.]',
      { italics: true, color: '808080' }) ] }),
].concat(
  phaseBlock('{{phase1_label}}', '{{phase1_comp}}'),
  phaseBlock('{{phase2_label}}', '{{phase2_comp}}'),
  phaseBlock('{{phase3_label}}', '{{phase3_comp}}'),
);

// ---------- PAGE 4: outlook + guidelines + disclaimer ----------
const page4 = [
  heading('Improve the route (optional but important)'),
  bullet([t('Staff: tailor optional next steps for this student (editable).')], GREEN),
  heading('After completing this path, you will be able to:'),
  bullet([t('Staff: list expected capabilities after the plan (editable).')], GREEN),
  heading('General Guidelines'),
  bullet([t('Commit to the plan while applying learning directly at work.')], GREEN),
  bullet([t('Focus on building practical skills, not only theoretical knowledge.')], GREEN),
  bullet([t('Track progress through measurable KPIs.')], GREEN),
  heading('Competitiveness Enhancement'),
  bullet([t('Develop strong negotiation and supplier management skills.')], GREEN),
  bullet([t('Use digital tools and AI in decision-making.')], GREEN),
  heading('Disclaimer'),
  new Paragraph({ shading: { type: ShadingType.CLEAR, fill: GREEN }, children: [ t(
    'This plan is based on professional development best practices and training programs provided ' +
    'by the center. Participation in these programs is not mandatory for competency development, and ' +
    'outcomes depend on individual performance, experience, and market conditions.') ] }),
];

// ---------- assemble ----------
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
