/**
 * Central configuration for the IDP student-report automation.
 *
 * BEFORE GOING LIVE, fill in the three IDs/email marked `FILL_ME`:
 *   - TEMPLATE_DOC_ID   : the Google Doc template you build from the reference report
 *   - OUTPUT_FOLDER_ID  : the Drive folder where generated drafts are saved
 *   - STAFF_EMAIL       : who gets the "draft ready for review" notification
 */

const CONFIG = {
  // Master data sheet (training programs catalogue).
  MASTER_SHEET_ID: '1mW1b8AB5Hx7DPH7BppnMit2kc9zQhHNkOuXiq1HtPHU',
  // Tab name inside the master sheet. Adjust if the tab is renamed.
  MASTER_TAB_NAME: 'Training Programs',
  // The master sheet has two decorative banner rows; the real header is row 3,
  // so program data begins on row 4.
  MASTER_HEADER_ROW: 3,

  // Google Doc template (built from the reference report).
  TEMPLATE_DOC_ID: '12a2wLrnAMH6zRWQwtp7d8STkPL_sQHFgyAhiZ8MSOlc',
  // Drive folder "IDP Reports 2026" for generated draft reports.
  OUTPUT_FOLDER_ID: '1WQV3wiu6PCMzlZPZKbppG7TjI607U1Cr',
  // Staff reviewer notified when a draft is ready.
  STAFF_EMAIL: 'info@elite-t-s.com',

  // Columns appended to the response sheet for pipeline bookkeeping.
  STATUS_COLUMN_HEADER: 'Status',
  REPORT_URL_COLUMN_HEADER: 'Report URL',
};

/**
 * Master sheet column headers, exactly as they appear on the header row.
 * Used to locate columns by name rather than fixed index.
 */
const MASTER_COLUMNS = {
  NAME: 'Training Program Name',
  RECOMMENDED_TYPE: 'Recomended Type', // (sic) matches the sheet's spelling
  DURATION: 'Duration',
  LEVEL: 'Level',
  LANGUAGE: 'Language',
  SPECIALITY: 'Speciality',
  SERVICE_PROVIDER: 'Service provider',
  PROGRAM_TYPE: 'Program type',
  COMPETENCIES: 'Competencies ( goal )',
  KPI: 'Performance indicator',
};

/**
 * Canonical student field -> a substring that uniquely identifies the form
 * response column header. Form headers are bilingual ("English | عربي"), so we
 * match by an English substring rather than exact equality.
 */
const FORM_FIELDS = {
  fullName: 'Full Name',
  email: 'Email address',
  age: 'Age',
  universityMajor: 'University major',
  currentField: 'Current Field',
  company: 'Company/Organization Name',
  jobTitle: 'Job Title',
  experienceYears: 'Years of Experience',
  field: 'field you want to develop in',
  preferredLanguage: 'Preferred language',
  englishLevel: 'English language level',
  mainGoal: 'Main Goal',
  shortGoals: 'Short-term Goals',
  longGoals: 'Long-term Goals',
  strengths: 'Strengths',
  placementResult: 'Placement test result',
  weaknesses: 'Weaknesses',
  currentChallenges: 'Current Challenges',
  currentSkills: 'Current Skills',
  skillsToDevelop: 'Skills to Develop',
  weeklyTime: 'Weekly Time Commitment',
  constraints: 'Any Constraints',
  learningStyle: 'Preferred Learning Style',
  budget: 'Do you have a budget',
  toolsUsed: 'Tools / Platforms Used',
  measureSuccess: 'How will you measure success',
  supportType: 'Preferred Support Type',
  notes: 'Additional Notes',
};

/**
 * Maps the free-text "field you want to develop in" answer to a canonical
 * master-sheet Speciality value. Keys are lowercased substrings (English or
 * Arabic) that may appear in the student's answer.
 */
const SPECIALITY_SYNONYMS = {
  'purchas': 'purchases',
  'procure': 'purchases',
  'مشتري': 'purchases',
  'logistic': 'Logistics',
  'لوجست': 'Logistics',
  'warehous': 'warehouses',
  'مستودع': 'warehouses',
  'مخازن': 'warehouses',
  'supply chain': 'Supply Chain',
  'سلسلة': 'Supply Chain',
  'الامداد': 'Supply Chain',
  'tourism': 'Tourism',
  'سياح': 'Tourism',
};

/**
 * Normalizes the many Level vocabularies in the master sheet into 3 ordered
 * tiers. tier 1 = foundational, 2 = intermediate, 3 = advanced.
 */
const LEVEL_TIERS = {
  'founding': 1,
  'beginner': 1,
  'medium': 2,
  'intermediate': 2,
  'advanced': 3,
};

/**
 * Phase definitions for the recommended-programs section, in display order.
 * `types` lists the master `Program type` values that fall into each phase.
 */
const PROGRAM_PHASES = [
  { title: 'Foundational Training Programs', types: ['Training course'] },
  { title: 'Professional Certifications', types: ['Professional Certificate'] },
  {
    title: 'Professional Qualifications',
    types: ['Professional qualification', 'International Certificate'],
  },
  { title: 'Workshops', types: ['workshop'] },
];

/**
 * Page-3 timeline stages, in display order, keyed to the master sheet's
 * `Recommended Type` values. `label` is what appears in the report callout.
 */
const TIMELINE_PHASES = [
  { recommendedType: 'Professional training phase', label: 'Professional training phase' },
  { recommendedType: 'Professional Skills Development Phase', label: 'Professional Skills Development Phase' },
  { recommendedType: 'Professional phase', label: 'Professional phase' },
];

/** Max competencies listed per timeline phase callout. */
CONFIG.MAX_PHASE_COMPETENCIES = 4;
