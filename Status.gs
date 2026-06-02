/**
 * Writes pipeline status back to the response sheet.
 *
 * Adds (once) the Status and Report URL columns if missing, then records the
 * outcome for the submitted row.
 */

/**
 * @param {Sheet}  sheet   the response sheet.
 * @param {number} rowNum  1-based row number of the submission.
 * @param {string} status  e.g. 'Draft Ready' or 'Error'.
 * @param {string} reportUrl optional URL of the generated draft.
 */
function recordStatus(sheet, rowNum, status, reportUrl) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = ensureColumn_(sheet, headers, CONFIG.STATUS_COLUMN_HEADER);
  const urlCol = ensureColumn_(sheet, headers, CONFIG.REPORT_URL_COLUMN_HEADER);

  sheet.getRange(rowNum, statusCol).setValue(status);
  if (reportUrl) {
    sheet.getRange(rowNum, urlCol).setValue(reportUrl);
  }
}

/**
 * Returns the 1-based column index for `header`, creating it at the end of the
 * header row if it doesn't exist yet.
 */
function ensureColumn_(sheet, headers, header) {
  const idx = headers.indexOf(header);
  if (idx !== -1) return idx + 1;
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue(header);
  headers.push(header); // keep local copy in sync for subsequent lookups
  return newCol;
}
