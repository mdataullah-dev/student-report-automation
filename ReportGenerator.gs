/**
 * Generates a draft report Doc from the template for one student.
 */

/**
 * Copies the template, fills placeholders + the programs table, and returns
 * the new Doc's URL.
 * @param {Object} student   normalized student object (see StudentData.gs).
 * @param {Object} matchResult output of matchPrograms().
 * @return {string} URL of the generated draft Doc.
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

  doc.saveAndClose();
  return newFile.getUrl();
}

/** Replaces all the simple {{token}} fields throughout the document body. */
function fillTextPlaceholders_(body, student, matchResult, dateStr) {
  const timeline = (matchResult && matchResult.timelinePhases) || [];
  const department = (matchResult && matchResult.speciality) || student.field;

  const map = {
    // Cover (page 1).
    '{{report_date}}': dateStr,
    '{{student_name}}': student.fullName,
    // Profile row (page 2).
    '{{department}}': department,
    '{{job_title}}': student.jobTitle,
    '{{experience}}': student.experienceYears,
    '{{experience_years}}': student.experienceYears,
    '{{field}}': student.field,
    '{{company}}': student.company,
    '{{preferred_language}}': student.preferredLanguage,
    '{{english_level}}': student.englishLevel,
    // Timeline competency callouts (page 3), derived from matched programs.
    '{{phase1_comp}}': phaseCompetencyText_(timeline[0]),
    '{{phase2_comp}}': phaseCompetencyText_(timeline[1]),
    '{{phase3_comp}}': phaseCompetencyText_(timeline[2]),
    '{{phase1_label}}': timeline[0] ? timeline[0].label : '',
    '{{phase2_label}}': timeline[1] ? timeline[1].label : '',
    '{{phase3_label}}': timeline[2] ? timeline[2].label : '',
    // Other optional profile fields, available if the template uses them.
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

/** Renders a timeline phase's competencies as newline-separated lines. */
function phaseCompetencyText_(phase) {
  if (!phase || !phase.competencies || phase.competencies.length === 0) return '';
  return phase.competencies.join('\n');
}

/** body.replaceText treats its first arg as regex; escape literal braces. */
function escapeForReplace_(token) {
  return token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Expands the programs table: finds the template body row (the one containing
 * {{prog_name}}), and for each phase inserts a phase-header row followed by one
 * row per program, then removes the template row.
 */
function fillProgramsTable_(body, matchResult) {
  const located = findTemplateRow_(body);
  if (!located) {
    return; // template has no programs table token; nothing to do
  }
  const table = located.table;
  const templateRowIndex = located.rowIndex;
  const templateRow = table.getRow(templateRowIndex);

  // Number of columns in the template row drives how we fill cells.
  const numCols = templateRow.getNumCells();

  if (matchResult.total === 0) {
    setRowCells_(templateRow, ['No matching programs found — please review and add manually.'], numCols);
    return;
  }

  let insertAt = templateRowIndex; // insert before the template row
  let counter = 0;

  matchResult.phases.forEach(function (phase) {
    // Phase sub-header row.
    const headerRow = table.insertTableRow(insertAt++, templateRow.copy());
    setRowCells_(headerRow, [phase.title], numCols);
    styleAsPhaseHeader_(headerRow);

    phase.programs.forEach(function (p) {
      counter++;
      const row = table.insertTableRow(insertAt++, templateRow.copy());
      // Column order matches the reference table: #, Course, Duration, Type, Action.
      const cells = buildProgramCells_(numCols, counter, p);
      setRowCells_(row, cells, numCols);
    });
  });

  // Remove the original template row (now after all inserted rows).
  table.removeRow(insertAt);
}

/**
 * Builds the cell values for one program row, tolerant of how many columns the
 * template table actually has (4 cols = no number column, 5 cols = numbered).
 */
function buildProgramCells_(numCols, index, p) {
  if (numCols >= 5) {
    return [String(index), p.name, p.duration, p.type, p.action];
  }
  // 4-column fallback: Course | Duration | Type | Action
  return [p.name, p.duration, p.type, p.action];
}

/** Writes an array of strings into a row's cells (pads/truncates to numCols). */
function setRowCells_(row, values, numCols) {
  for (let c = 0; c < numCols; c++) {
    const val = c < values.length ? String(values[c] == null ? '' : values[c]) : '';
    row.getCell(c).setText(val);
  }
}

/** Visually distinguishes a phase header row (bold, light background). */
function styleAsPhaseHeader_(row) {
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    cell.setBackgroundColor('#D9E1F2');
    cell.editAsText().setBold(true);
  }
}

/**
 * Finds the table and row index containing the {{prog_name}} token.
 * @return {{table: Table, rowIndex: number}|null}
 */
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
