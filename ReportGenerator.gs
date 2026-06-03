/**
 * Generates a draft report Doc from the template for one student.
 *
 * Pipeline: copy template -> fill inline tokens -> expand programs table ->
 * insert real bulleted lists at list-tokens -> strip all highlights -> save.
 */

function generateReport(student, matchResult) {
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);

  const name = student.fullName || 'Unknown Student';
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const safeName = name.replace(/[\\/:*?"<>|]/g, ' ').substring(0, 80).trim();
  const copyName = 'Report - ' + safeName + ' - ' + dateStr;

  const newFile = templateFile.makeCopy(copyName, folder);
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  fillTextPlaceholders_(body, student, matchResult, dateStr);
  fillProgramsTable_(body, matchResult);
  fillBulletLists_(body, student, matchResult);
  clearHighlights_(body);

  doc.saveAndClose();
  return newFile.getUrl();
}

/* ------------------------------------------------------------------ */
/* Inline {{token}} replacement                                        */
/* ------------------------------------------------------------------ */

/** Replaces the simple inline {{token}} fields throughout the body. */
function fillTextPlaceholders_(body, student, matchResult, dateStr) {
  const timeline = (matchResult && matchResult.timelinePhases) || [];
  const department = titleCase_(englishSide_((matchResult && matchResult.speciality) || student.field));

  const map = {
    // Cover (page 1).
    '{{report_date}}': dateStr,
    '{{student_name}}': student.fullName,
    // Profile row (page 2) — cleaned to the English side where bilingual.
    '{{department}}': department,
    '{{job_title}}': student.jobTitle,
    '{{experience}}': formatExperience_(student.experienceYears),
    '{{experience_years}}': student.experienceYears,
    '{{field}}': englishSide_(student.field),
    '{{company}}': student.company,
    '{{preferred_language}}': englishSide_(student.preferredLanguage),
    '{{english_level}}': student.englishLevel,
    // Timeline phase headings (page 3); the bullets are filled separately.
    '{{phase1_label}}': timeline[0] ? timeline[0].label : '',
    '{{phase2_label}}': timeline[1] ? timeline[1].label : '',
    '{{phase3_label}}': timeline[2] ? timeline[2].label : '',
    // Optional profile fields, if the template references them.
    '{{main_goal}}': student.mainGoal,
    '{{short_goals}}': student.shortGoals,
    '{{long_goals}}': student.longGoals,
    '{{strengths}}': student.strengths,
    '{{weaknesses}}': student.weaknesses,
    '{{skills_to_develop}}': student.skillsToDevelop,
  };
  Object.keys(map).forEach(function (token) {
    body.replaceText(escapeForReplace_(token), map[token] == null ? '' : String(map[token]));
  });
}

/** body.replaceText / findText treat the arg as regex; escape literal braces. */
function escapeForReplace_(token) {
  return token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** "AR | EN" -> the English (latin) side; plain values pass through. */
function englishSide_(v) {
  if (!v) return '';
  const s = String(v);
  if (s.indexOf('|') === -1) return s.trim();
  const parts = s.split('|').map(function (p) { return p.trim(); });
  for (let i = 0; i < parts.length; i++) {
    if (/[A-Za-z]/.test(parts[i])) return parts[i];
  }
  return parts[parts.length - 1];
}

/** Title-cases a short phrase: "warehouses" -> "Warehouses". */
function titleCase_(s) {
  if (!s) return '';
  return String(s).replace(/\w\S*/g, function (w) {
    return w.charAt(0).toUpperCase() + w.substr(1);
  });
}

/** "6" -> "6 years"; passes through values that already include text. */
function formatExperience_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  return /^\d+$/.test(s) ? s + ' years' : s;
}

/* ------------------------------------------------------------------ */
/* Bulleted-list sections                                              */
/* ------------------------------------------------------------------ */

/** Fills every list-token paragraph with real bullet items. */
function fillBulletLists_(body, student, matchResult) {
  const tl = matchResult.timelinePhases || [];
  const phaseFallback = 'To be confirmed based on the selected programs.';

  replaceParagraphWithBullets_(body, '{{skills_list}}', generateSkills(student, matchResult));
  replaceParagraphWithBullets_(body, '{{phase1_list}}', (tl[0] && tl[0].competencies) || [], phaseFallback);
  replaceParagraphWithBullets_(body, '{{phase2_list}}', (tl[1] && tl[1].competencies) || [], phaseFallback);
  replaceParagraphWithBullets_(body, '{{phase3_list}}', (tl[2] && tl[2].competencies) || [], phaseFallback);
  replaceParagraphWithBullets_(body, '{{improve_list}}', generateImproveRoute(student, matchResult));
  replaceParagraphWithBullets_(body, '{{outcome_list}}', generateOutcomes(student, matchResult));
}

/**
 * Finds the paragraph containing `token` and replaces it with a real bulleted
 * list of `items`. If items is empty, uses `fallback` (or removes the line).
 */
function replaceParagraphWithBullets_(body, token, items, fallback) {
  const range = body.findText(escapeForReplace_(token));
  if (!range) return;

  let para = range.getElement();
  while (para && para.getParent && para.getType() !== DocumentApp.ElementType.PARAGRAPH &&
         para.getType() !== DocumentApp.ElementType.LIST_ITEM) {
    para = para.getParent();
  }
  if (!para) return;

  const parent = para.getParent();
  const idx = parent.getChildIndex(para);

  let list = items && items.length ? items.slice() : (fallback ? [fallback] : []);
  if (list.length === 0) { parent.removeChild(para); return; }

  for (let i = 0; i < list.length; i++) {
    const li = parent.insertListItem(idx + i, String(list[i]));
    li.setGlyphType(DocumentApp.GlyphType.BULLET);
    li.setNestingLevel(0);
  }
  // Remove the original token paragraph (now pushed below the inserted items).
  parent.removeChild(parent.getChild(idx + list.length));
}

/* ------------------------------------------------------------------ */
/* Highlight stripping                                                 */
/* ------------------------------------------------------------------ */

/** Recursively clears every text-background highlight in the document. */
function clearHighlights_(el) {
  const type = el.getType && el.getType();
  if (type === DocumentApp.ElementType.TEXT) {
    const t = el.asText();
    const len = t.getText().length;
    if (len > 0) t.setBackgroundColor(0, len - 1, null);
    return;
  }
  if (el.getNumChildren) {
    for (let i = 0; i < el.getNumChildren(); i++) {
      clearHighlights_(el.getChild(i));
    }
  }
}

/* ------------------------------------------------------------------ */
/* Programs table expansion                                            */
/* ------------------------------------------------------------------ */

function fillProgramsTable_(body, matchResult) {
  const located = findTemplateRow_(body);
  if (!located) return;

  const table = located.table;
  const templateRowIndex = located.rowIndex;
  const templateRow = table.getRow(templateRowIndex);
  const numCols = templateRow.getNumCells();

  if (matchResult.total === 0) {
    setRowCells_(templateRow, ['No matching programs found — please review and add manually.'], numCols);
    return;
  }

  // Flat, continuously-numbered list like the reference report (no phase
  // sub-header rows — the phased view lives on the page-3 timeline).
  let insertAt = templateRowIndex;
  let counter = 0;

  matchResult.phases.forEach(function (phase) {
    phase.programs.forEach(function (p) {
      counter++;
      const row = table.insertTableRow(insertAt++, templateRow.copy());
      setRowCells_(row, buildProgramCells_(numCols, counter, p), numCols);
    });
  });

  table.removeRow(insertAt);
}

function buildProgramCells_(numCols, index, p) {
  if (numCols >= 5) {
    return [String(index), p.name, p.duration, p.type, p.action];
  }
  return [p.name, p.duration, p.type, p.action];
}

function setRowCells_(row, values, numCols) {
  for (let c = 0; c < numCols; c++) {
    const val = c < values.length ? String(values[c] == null ? '' : values[c]) : '';
    row.getCell(c).setText(val);
  }
}

function findTemplateRow_(body) {
  const tables = body.getTables();
  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    for (let r = 0; r < table.getNumRows(); r++) {
      if (table.getRow(r).getText().indexOf('{{prog_name}}') !== -1) {
        return { table: table, rowIndex: r };
      }
    }
  }
  return null;
}
