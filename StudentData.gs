/**
 * Reading and normalizing a student's form response into a plain object.
 */

/**
 * Finds the index of a header that contains the given English substring.
 * Form headers are bilingual ("English | عربي"), matched case-insensitively.
 * @return {number} zero-based index, or -1 if not found.
 */
function findHeaderIndex_(headers, substring) {
  const needle = substring.toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().indexOf(needle) !== -1) {
      return i;
    }
  }
  return -1;
}

/**
 * Builds a normalized student object from a response row.
 * @param {Array} row     the submitted values for one response.
 * @param {Array} headers the header row of the response sheet.
 * @return {Object} keyed by the canonical names in FORM_FIELDS.
 */
function readStudent(row, headers) {
  const student = {};
  Object.keys(FORM_FIELDS).forEach(function (key) {
    const idx = findHeaderIndex_(headers, FORM_FIELDS[key]);
    const raw = idx === -1 ? '' : row[idx];
    student[key] = raw === null || raw === undefined ? '' : String(raw).trim();
  });

  // Coerce years of experience to a number (pull the first numeric token).
  student.experienceYearsNum = parseExperienceYears_(student.experienceYears);
  return student;
}

/**
 * Extracts a numeric year count from a free-text experience answer.
 * Handles "6", "6 years", "أكثر من 8", "8+" etc. Defaults to 0.
 */
function parseExperienceYears_(value) {
  const match = String(value).match(/\d+/);
  return match ? parseInt(match[0], 10) : 0;
}
