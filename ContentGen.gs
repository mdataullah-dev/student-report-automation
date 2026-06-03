/**
 * Rule-based generation of the narrative bullet sections, so the report reads
 * like the reference without needing an LLM. Each function returns an array of
 * short bullet strings. Content is keyed by the student's matched speciality and
 * blended with the competencies of the actually-matched programs.
 */

/** Per-speciality content library. `generic` is the fallback. */
const SPECIALITY_CONTENT = {
  purchases: {
    label: 'procurement',
    skills: [
      'Strong understanding of procurement strategy and the purchasing cycle',
      'Ability to select, evaluate, and manage suppliers professionally',
      'Improved cost control, negotiation, and sourcing decisions',
      'Contract management and procurement governance awareness',
    ],
    improve: [
      'Learn ERP procurement modules such as SAP (MM) or Oracle',
      'Develop advanced spend-analysis skills (Excel / Power BI)',
      'Pursue advanced certifications later (e.g., CIPS Level 4+)',
      'Build a professional procurement portfolio of delivered savings',
      'Expand your professional network in the procurement field',
    ],
    outcomes: [
      'Transition confidently into procurement and sourcing roles',
      'Manage procurement processes and supplier relationships end-to-end',
      'Conduct negotiations and manage contracts effectively',
      'Analyze costs and improve procurement efficiency',
      'Support strategic sourcing and category decisions',
    ],
  },
  warehouses: {
    label: 'warehousing & inventory',
    skills: [
      'Solid grasp of warehouse and inventory operations management',
      'Ability to use warehouse management systems (WMS) effectively',
      'Improved receiving, storage, and distribution processes',
      'Inventory accuracy, governance, and space optimization',
    ],
    improve: [
      'Gain hands-on experience with a WMS / ERP inventory module',
      'Develop data-analysis skills for stock and demand reporting',
      'Pursue advanced warehousing certifications (e.g., Level 5)',
      'Adopt global best practices in warehouse layout and safety',
      'Build cross-functional links with procurement and logistics',
    ],
    outcomes: [
      'Oversee warehouse operations with confidence',
      'Improve internal storage and distribution efficiency',
      'Apply global best practices in warehouse management',
      'Integrate warehousing with procurement and logistics',
      'Make data-driven inventory and capacity decisions',
    ],
  },
  Logistics: {
    label: 'logistics & transport',
    skills: [
      'Understanding of transport, distribution, and shipping operations',
      'Ability to optimize routes and reduce logistics costs',
      'Knowledge of international trade and customs processes',
      'Improved on-time delivery and service-level performance',
    ],
    improve: [
      'Learn transport/logistics planning and TMS tools',
      'Develop cost-per-shipment and lead-time analytics skills',
      'Pursue advanced logistics diplomas (e.g., Level 5)',
      'Study international shipping regulations and compliance',
      'Build relationships with carriers and 3PL partners',
    ],
    outcomes: [
      'Plan and manage transport and distribution operations',
      'Optimize logistics networks and reduce costs',
      'Manage international shipping and compliance',
      'Improve delivery reliability and service levels',
      'Make data-driven logistics decisions',
    ],
  },
  'Supply Chain': {
    label: 'supply chain',
    skills: [
      'End-to-end understanding of supply chain operations',
      'Ability to align supply chain activities with company strategy',
      'Demand planning, inventory, and risk-management skills',
      'Data-driven decision-making across the supply chain',
    ],
    improve: [
      'Develop S&OP and demand-planning capabilities',
      'Strengthen data analytics (Excel / Power BI / ERP)',
      'Pursue advanced supply chain certifications (e.g., CSCP)',
      'Study supply chain resilience and risk management',
      'Build a strategic, cross-functional professional profile',
    ],
    outcomes: [
      'Manage integrated supply chain operations strategically',
      'Align planning, procurement, and warehousing',
      'Use ERP/MRP systems to improve efficiency',
      'Make data-driven, end-to-end decisions',
      'Support supply chain strategy and resilience',
    ],
  },
  Tourism: {
    label: 'tourism',
    skills: [
      'Understanding of destination and tourism product management',
      'Service quality and visitor-experience management skills',
      'Digital tourism marketing and branding awareness',
      'Sustainable tourism and heritage-preservation principles',
    ],
    improve: [
      'Develop digital tourism marketing skills',
      'Study destination strategic planning and investment',
      'Pursue advanced tourism certifications',
      'Learn service-quality and hospitality standards',
      'Build a professional network in the tourism sector',
    ],
    outcomes: [
      'Manage tourism products and destinations',
      'Improve visitor experience and service quality',
      'Lead digital tourism marketing initiatives',
      'Apply sustainable tourism practices',
      'Support tourism strategy and investment decisions',
    ],
  },
  generic: {
    label: 'your chosen field',
    skills: [
      'Strong foundation in core concepts and best practices',
      'Practical, workplace-ready professional skills',
      'Improved analytical and decision-making ability',
      'Effective communication with stakeholders',
    ],
    improve: [
      'Apply learning directly at work to reinforce skills',
      'Develop relevant digital and analytical tools',
      'Pursue advanced certifications over time',
      'Build a strong professional portfolio',
      'Expand your professional network in the field',
    ],
    outcomes: [
      'Apply new skills confidently in your role',
      'Take on broader responsibilities in your field',
      'Make better, data-informed decisions',
      'Strengthen your professional profile',
    ],
  },
};

/** Returns the content block for the student's matched speciality. */
function contentFor_(matchResult) {
  const key = matchResult && matchResult.speciality;
  if (key && SPECIALITY_CONTENT[key]) return SPECIALITY_CONTENT[key];
  // case-insensitive fallback match
  if (key) {
    const lower = String(key).toLowerCase();
    const found = Object.keys(SPECIALITY_CONTENT).filter(function (k) {
      return k.toLowerCase() === lower;
    })[0];
    if (found) return SPECIALITY_CONTENT[found];
  }
  return SPECIALITY_CONTENT.generic;
}

/**
 * "The skills that will be acquired" — 3-4 bullets. Prefers competencies drawn
 * from the student's actually-matched programs; tops up from the speciality
 * library if there aren't enough.
 */
function generateSkills(student, matchResult) {
  const block = contentFor_(matchResult);
  const fromCourses = [];
  const seen = {};
  (matchResult.timelinePhases || []).forEach(function (ph) {
    (ph.competencies || []).forEach(function (c) {
      const key = c.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      fromCourses.push(c);
    });
  });
  const out = fromCourses.slice(0, 4);
  for (let i = 0; out.length < 3 && i < block.skills.length; i++) {
    if (out.indexOf(block.skills[i]) === -1) out.push(block.skills[i]);
  }
  return out.slice(0, 4);
}

/** "Improve the route (optional but important)" — 4-5 bullets. */
function generateImproveRoute(student, matchResult) {
  return contentFor_(matchResult).improve.slice(0, 5);
}

/** "After completing this path, you will be able to" — 4-5 bullets. */
function generateOutcomes(student, matchResult) {
  return contentFor_(matchResult).outcomes.slice(0, 5);
}
