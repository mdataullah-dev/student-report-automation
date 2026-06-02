/**
 * Master-data reading and the program-matching logic.
 *
 * matchPrograms() is intentionally pure (no Sheet/Doc calls) so it can be
 * unit-tested from the Apps Script editor with a hand-built student object.
 */

/**
 * Reads the master sheet into an array of plain row objects keyed by the
 * MASTER_COLUMNS values. Skips the decorative banner rows above the header.
 * @return {Array<Object>}
 */
function readMasterRows() {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MASTER_TAB_NAME) || ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();

  const headerIdx = CONFIG.MASTER_HEADER_ROW - 1; // to zero-based
  const headers = values[headerIdx].map(function (h) { return String(h).trim(); });

  const colIndex = {};
  Object.keys(MASTER_COLUMNS).forEach(function (key) {
    colIndex[key] = headers.indexOf(MASTER_COLUMNS[key]);
  });

  const rows = [];
  for (let r = headerIdx + 1; r < values.length; r++) {
    const row = values[r];
    const name = colIndex.NAME >= 0 ? String(row[colIndex.NAME]).trim() : '';
    if (!name) continue; // skip blank/spacer rows
    rows.push({
      name: name,
      recommendedType: getCell_(row, colIndex.RECOMMENDED_TYPE),
      duration: getCell_(row, colIndex.DURATION),
      level: getCell_(row, colIndex.LEVEL),
      language: getCell_(row, colIndex.LANGUAGE),
      speciality: getCell_(row, colIndex.SPECIALITY),
      serviceProvider: getCell_(row, colIndex.SERVICE_PROVIDER),
      programType: getCell_(row, colIndex.PROGRAM_TYPE),
      competencies: getCell_(row, colIndex.COMPETENCIES),
      kpi: getCell_(row, colIndex.KPI),
    });
  }
  return rows;
}

function getCell_(row, idx) {
  return idx >= 0 && row[idx] != null ? String(row[idx]).trim() : '';
}

/**
 * Maps the student's free-text "field to develop" to a canonical Speciality.
 * @return {string} canonical speciality, or '' if unrecognized.
 */
function resolveSpeciality_(fieldText) {
  const text = String(fieldText).toLowerCase();
  const keys = Object.keys(SPECIALITY_SYNONYMS);
  for (let i = 0; i < keys.length; i++) {
    if (text.indexOf(keys[i]) !== -1) return SPECIALITY_SYNONYMS[keys[i]];
  }
  return '';
}

/** Normalized level vocabulary -> tier number (1..3), or 0 if unknown. */
function levelTier_(level) {
  return LEVEL_TIERS[String(level).trim().toLowerCase()] || 0;
}

/** Experience years -> target tier (1..3). */
function experienceTier_(years) {
  if (years <= 2) return 1;
  if (years <= 7) return 2;
  return 3;
}

/**
 * True if a program's language is compatible with the student's preference.
 * "Arabic & English" programs are always compatible.
 */
function languageCompatible_(programLanguage, preferred) {
  const prog = String(programLanguage).toLowerCase();
  const pref = String(preferred).toLowerCase();
  if (prog.indexOf('&') !== -1 || (prog.indexOf('arabic') !== -1 && prog.indexOf('english') !== -1)) {
    return true; // bilingual program suits anyone
  }
  const prefersArabic = pref.indexOf('arabic') !== -1 || pref.indexOf('عرب') !== -1;
  const prefersEnglish = pref.indexOf('english') !== -1 || pref.indexOf('انجل') !== -1 || pref.indexOf('إنجل') !== -1;
  if (prefersArabic && prefersEnglish) return true;
  if (prefersArabic) return prog.indexOf('arabic') !== -1;
  if (prefersEnglish) return prog.indexOf('english') !== -1;
  return true; // no clear preference -> don't exclude
}

/**
 * Splits a Competencies (goal) cell into individual items. The master data is
 * inconsistent: some cells are newline-separated lists, others are a single
 * comma-separated line. Prefer newlines; only fall back to commas when the cell
 * is a single line (so comma-containing phrases aren't chopped).
 */
function splitCompetencies_(text) {
  if (!text) return [];
  let parts = String(text).split(/[\r\n]+/).map(function (s) { return s.trim(); })
    .filter(function (s) { return s; });
  if (parts.length <= 1) {
    parts = String(text).split(',').map(function (s) { return s.trim(); })
      .filter(function (s) { return s; });
  }
  return parts;
}

/**
 * Short "Action" description from the competencies field: first item only.
 */
function firstCompetency_(competencies) {
  const parts = splitCompetencies_(competencies);
  return parts.length ? parts[0] : '';
}

/**
 * Core matching. Pure function over a student object and master rows.
 * @return {{speciality:string, targetTier:number, phases:Array, total:number}}
 *   phases: [{ title, programs:[{name,duration,type,action}] }]
 */
function matchPrograms(student, masterRows) {
  const speciality = resolveSpeciality_(student.field);
  const targetTier = experienceTier_(student.experienceYearsNum || 0);

  // 1+2+3: filter by speciality, language, and level tier (target and below).
  const candidates = masterRows.filter(function (row) {
    if (speciality && row.speciality.toLowerCase() !== speciality.toLowerCase()) return false;
    if (!languageCompatible_(row.language, student.preferredLanguage)) return false;
    const tier = levelTier_(row.level);
    // Keep unknown-tier rows (tier 0) so nothing useful is silently dropped,
    // plus any row at or below the student's target tier.
    return tier === 0 || tier <= targetTier;
  });

  // 4: group into phases by Program type, ordered by tier then de-duplicated.
  const phases = PROGRAM_PHASES.map(function (phaseDef) {
    const matches = candidates
      .filter(function (row) {
        return phaseDef.types.some(function (t) {
          return row.programType.toLowerCase() === t.toLowerCase();
        });
      })
      .sort(function (a, b) { return levelTier_(a.level) - levelTier_(b.level); });

    const seen = {};
    const programs = [];
    matches.forEach(function (row) {
      const key = row.name.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      programs.push({
        name: row.name,
        duration: row.duration,
        type: row.recommendedType || row.programType,
        action: firstCompetency_(row.competencies),
      });
    });
    return { title: phaseDef.title, programs: programs };
  }).filter(function (phase) { return phase.programs.length > 0; });

  const total = phases.reduce(function (sum, p) { return sum + p.programs.length; }, 0);

  // Page-3 timeline: group the same candidates by master "Recommended Type"
  // into the 3 development stages, aggregating their competencies.
  const timelinePhases = buildTimelinePhases_(candidates);

  return {
    speciality: speciality,
    targetTier: targetTier,
    phases: phases,
    total: total,
    timelinePhases: timelinePhases,
  };
}

/**
 * Builds the 3 timeline stages (in display order) from candidate rows, each
 * with a de-duplicated list of competencies pulled from its programs.
 * @return {Array<{label:string, competencies:Array<string>}>}
 */
function buildTimelinePhases_(candidates) {
  return TIMELINE_PHASES.map(function (stage) {
    const inStage = candidates.filter(function (row) {
      return row.recommendedType.trim().toLowerCase() === stage.recommendedType.toLowerCase();
    });

    const seen = {};
    const competencies = [];
    inStage.forEach(function (row) {
      splitCompetencies_(row.competencies).forEach(function (comp) {
        const key = comp.toLowerCase();
        if (seen[key]) return;
        seen[key] = true;
        competencies.push(comp);
      });
    });

    return {
      label: stage.label,
      competencies: competencies.slice(0, CONFIG.MAX_PHASE_COMPETENCIES),
    };
  });
}

/**
 * Editor-run helper: smoke-test the matcher without a form submission.
 * Adjust the fake student and check the logs.
 */
function testMatcher() {
  const fake = {
    field: 'warehouses',
    preferredLanguage: 'English',
    experienceYearsNum: 6,
  };
  const result = matchPrograms(fake, readMasterRows());
  Logger.log('speciality=%s tier=%s total=%s', result.speciality, result.targetTier, result.total);
  result.phases.forEach(function (phase) {
    Logger.log('PHASE: %s (%s)', phase.title, phase.programs.length);
    phase.programs.forEach(function (p) {
      Logger.log('  - %s | %s | %s', p.name, p.duration, p.type);
    });
  });
  result.timelinePhases.forEach(function (tp) {
    Logger.log('TIMELINE: %s -> %s', tp.label, tp.competencies.join(' | '));
  });
}
