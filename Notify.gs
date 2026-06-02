/**
 * Staff notification for a freshly generated draft report.
 */

/**
 * Emails the staff reviewer a link to the draft. No report is sent to the
 * student here — staff send manually after reviewing/editing.
 * @param {string} docUrl  URL of the generated draft Doc.
 * @param {Object} student normalized student object.
 * @param {Object} matchResult output of matchPrograms() (for a quick summary).
 */
function notifyStaff(docUrl, student, matchResult) {
  const name = student.fullName || 'Unknown Student';
  const subject = 'IDP draft ready for review — ' + name;

  const lines = [
    'A new IDP report draft has been generated and is ready for review.',
    '',
    'Student: ' + name,
    'Email: ' + (student.email || '(none)'),
    'Company: ' + (student.company || '(none)'),
    'Job title: ' + (student.jobTitle || '(none)'),
    'Experience: ' + (student.experienceYears || '(none)'),
    'Field to develop: ' + (student.field || '(none)') +
      (matchResult && matchResult.speciality ? ' -> ' + matchResult.speciality : ''),
    'Preferred language: ' + (student.preferredLanguage || '(none)'),
    'Matched programs: ' + (matchResult ? matchResult.total : 'n/a'),
    '',
    'Review / edit the draft here:',
    docUrl,
    '',
    'Remember to edit the green sections (Guidelines, Competitiveness, Disclaimer)',
    'before sending the final report to the student.',
  ];

  MailApp.sendEmail(CONFIG.STAFF_EMAIL, subject, lines.join('\n'));
}
