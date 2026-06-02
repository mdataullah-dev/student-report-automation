/**
 * Pipeline entry points and trigger installation.
 */

/**
 * Run this ONCE from the Apps Script editor to register the installable
 * on-form-submit trigger. An installable trigger (not a simple one) is required
 * because the pipeline sends email and opens other files / uses Drive.
 */
function installTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Avoid duplicate triggers if re-run.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('Installed onFormSubmit trigger.');
}

/**
 * Main pipeline: runs on every form submission.
 * @param {Object} e form-submit event (has e.range, e.values, e.namedValues).
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const rowNum = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

  try {
    const student = readStudent(row, headers);
    const masterRows = readMasterRows();
    const matchResult = matchPrograms(student, masterRows);

    const docUrl = generateReport(student, matchResult);
    notifyStaff(docUrl, student, matchResult);
    recordStatus(sheet, rowNum, 'Draft Ready', docUrl);
  } catch (err) {
    recordStatus(sheet, rowNum, 'Error: ' + err.message, '');
    // Surface failures to staff so submissions aren't silently dropped.
    try {
      MailApp.sendEmail(
        CONFIG.STAFF_EMAIL,
        'IDP automation error (row ' + rowNum + ')',
        'The report pipeline failed for row ' + rowNum + ':\n\n' + err.stack
      );
    } catch (mailErr) {
      // Nothing more we can do; the status cell already records the failure.
    }
    throw err; // re-throw so it shows in the execution log too
  }
}

/**
 * Editor-run helper (no arguments — pick this in the Run dropdown): generates a
 * report for the LAST response row already in the sheet. Use this to test on an
 * existing submission without resubmitting the form.
 */
function testGenerateLatestRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No response rows found (row 1 is the header).');
    return;
  }
  runForRow(lastRow);
}

/**
 * Editor-run helper: re-run the pipeline for a single existing row (1-based),
 * useful for testing without resubmitting the form.
 */
function runForRow(rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

  const student = readStudent(row, headers);
  const matchResult = matchPrograms(student, readMasterRows());
  const docUrl = generateReport(student, matchResult);
  notifyStaff(docUrl, student, matchResult);
  recordStatus(sheet, rowNum, 'Draft Ready', docUrl);
  Logger.log('Generated: %s', docUrl);
}
