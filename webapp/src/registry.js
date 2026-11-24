/**
 * The converters the app may run. Only these scripts, with only these arguments, are ever started - nothing the browser sends becomes a command.
 * Every converter also takes --from-date, --to-date, --ref (an earlier upload CSV to compare with) and --out; the app adds those itself.
 */
export const COMMON_INPUTS = [
  { role: 'ref', flag: '--ref', label: 'Earlier upload to compare with (optional)', accept: ['.csv'], required: false },
];

export const SCRIPTS = [
  {
    id: 'liberty',
    insurer: 'Liberty General',
    label: 'Liberty - branch grid',
    file: 'liberty_grid_to_config.py',
    outPrefix: 'LIBERTY',
    inputs: [{ role: 'grid', flag: '--grid', label: 'Liberty branch grid', accept: ['.xlsx', '.xlsm'], required: true }],
    options: [],
    secrets: [],
  },
  {
    id: 'magma',
    insurer: 'Magma HDI',
    label: 'Magma - comprehensive + SATP grids',
    file: 'magma_grid_to_config.py',
    outPrefix: 'MAGMA',
    inputs: [
      { role: 'grid', flag: '--grid', label: 'Comprehensive grid', accept: ['.xlsx', '.xlsm'], required: true },
      { role: 'satp', flag: '--satp', label: 'SATP grid', accept: ['.xlsb', '.xlsx'], required: true },
    ],
    options: [],
    secrets: [],
  },
  {
    id: 'rsa_gcv4w',
    insurer: 'Royal Sundaram',
    label: 'Royal Sundaram - GCV 4W comprehensive',
    file: 'rsa_gcv4w_comp_to_config.py',
    outPrefix: 'RSA_GCV4W_COMP',
    inputs: [{ role: 'grid', flag: '--grid', label: 'RSA GCV 4W workbook', accept: ['.xlsx', '.xlsm'], required: true }],
    options: [{ name: 'prefix', flag: '--prefix', type: 'text', label: 'RTO group prefix (empty = the script default)', pattern: '^[A-Za-z0-9_\\-]{1,60}$' }],
    secrets: [],
  },
  {
    id: 'rsa_pvtcar',
    insurer: 'Royal Sundaram',
    label: 'Royal Sundaram - Private Car',
    file: 'rsa_pvtcar_to_config.py',
    outPrefix: 'RSA_PVTCAR',
    inputs: [{ role: 'grid', flag: '--grid', label: 'RSA Private Car workbook', accept: ['.xlsx', '.xlsm'], required: true }],
    options: [{ name: 'prefix', flag: '--prefix', type: 'text', label: 'RTO group prefix (empty = the script default)', pattern: '^[A-Za-z0-9_\\-]{1,60}$' }],
    secrets: [],
  },
  {
    id: 'united',
    insurer: 'United India',
    label: 'United India - commission circular',
    file: 'united_commission_to_config.py',
    outPrefix: 'UNITED',
    inputs: [{ role: 'grid', flag: '--grid', label: 'Circular (PDF, or its transcription .json)', accept: ['.pdf', '.json'], required: true }],
    options: [],
    secrets: [],
  },
  {
    id: 'shriram_broker',
    insurer: 'Shriram General',
    label: 'Shriram - broker grid',
    file: 'shriram_broker_grid_to_config.py',
    outPrefix: 'SHRIRAM',
    inputs: [
      { role: 'grid', flag: '--grid', label: 'Broker grid', accept: ['.xlsx', '.xlsm'], required: true },
      { role: 'addendum', flag: '--addendum', label: 'Addendum workbook (when Shriram sent one)', accept: ['.xlsx', '.xlsm'], required: false },
    ],
    options: [{ name: 'writeEv', flag: '--write-ev', type: 'flag', label: 'Also write Electric rows for Private Car and taxis' }],
    secrets: [],
  },
  {
    id: 'shriram_2w',
    insurer: 'Shriram General',
    label: 'Shriram - two-wheeler grid',
    file: 'shriram_2w_to_config.py',
    outPrefix: 'SHRIRAM_TW',
    inputs: [
      { role: 'grid', flag: '--grid', label: 'Two-wheeler workbook', accept: ['.xlsx', '.xlsm'], required: true },
      { role: 'brokerGrid', flag: '--broker-grid', label: 'Broker grid (its PAN India sheet; optional)', accept: ['.xlsx', '.xlsm'], required: false },
    ],
    options: [],
    secrets: [],
  },
  {
    id: 'godigit',
    insurer: 'Go Digit',
    label: 'Go Digit - agency grid',
    file: 'godigit_agency_to_config.py',
    outPrefix: 'GODIGIT',
    inputs: [{ role: 'grid', flag: '--grid', label: 'Agency workbook (password protected)', accept: ['.xls', '.xlsx'], required: true }],
    options: [],
    secrets: [{ name: 'password', env: 'GRID_PASSWORD', label: 'Workbook password (kept in memory for this job only, never saved)' }],
  },
  {
    id: 'sbi',
    insurer: 'SBI General',
    label: 'SBI General - broker grid',
    file: 'sbi_grid_to_config.py',
    outPrefix: 'SBI',
    inputs: [{ role: 'grid', flag: '--grid', label: 'SBI broker grid (Agency, C & D)', accept: ['.xlsx', '.xlsm'], required: true }],
    options: [],
    secrets: [],
  },
  {
    id: 'icici',
    insurer: 'ICICI Lombard',
    label: 'ICICI Lombard - CV, Private Car and Two-Wheeler grids',
    file: 'icici_grid_to_config.py',
    outPrefix: 'ICICI',
    inputs: [
      { role: 'cvGrid', flag: '--cv-grid', label: 'CV grid (COMP / AOTP / Council Call sheets) - choose at least one of the four grids', accept: ['.xlsx', '.xlsm'], required: false },
      { role: 'pcGrid', flag: '--pc-grid', label: 'Private Car grid', accept: ['.xlsx', '.xlsm'], required: false },
      { role: 'twGrid', flag: '--tw-grid', label: 'Two-Wheeler grid (TW Old)', accept: ['.xlsx', '.xlsm'], required: false },
      { role: 'twNewGrid', flag: '--tw-new-grid', label: 'New Two-Wheeler (1+5) grid, one sheet per OEM', accept: ['.xlsx', '.xlsm'], required: false },
    ],
    options: [],
    secrets: [],
  },
  {
    id: 'indusind',
    insurer: 'IndusInd General (Reliance)',
    label: 'IndusInd General (Reliance) - commercial-vehicle grid',
    file: 'indusind_grid_to_config.py',
    outPrefix: 'INDUSIND',
    inputs: [{ role: 'grid', flag: '--grid', label: 'IndusInd / Reliance TC grid for brokers', accept: ['.xlsx', '.xlsm'], required: true }],
    options: [],
    secrets: [],
  },
];

const byId = new Map(SCRIPTS.map((s) => [s.id, s]));

export function getScript(id) {
  return byId.get(id) || null;
}

/** What the browser needs to build the upload form. */
export function publicScripts() {
  return SCRIPTS.map((s) => ({
    id: s.id, insurer: s.insurer, label: s.label,
    inputs: [...s.inputs, ...COMMON_INPUTS].map(({ role, label, accept, required }) => ({ role, label, accept, required })),
    options: s.options.map(({ name, type, label }) => ({ name, type, label })),
    secrets: s.secrets.map(({ name, label }) => ({ name, label })),
  }));
}

export function outputName(script, fromDate) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(fromDate);
  return `${script.outPrefix}_UPLOAD_${m[3]}-${m[2]}.csv`;
}

/** DD/MM/YYYY that is a real date. */
export function validDate(text) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(text || ''));
  if (!m) return false;
  const d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  return d.getUTCFullYear() === +m[3] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[1];
}
