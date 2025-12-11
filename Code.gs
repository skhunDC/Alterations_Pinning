const SPREADSHEET_NAME = 'Dublin Cleaners Training';
const PROGRESS_SHEET_NAME = 'TrainingProgress';
const HEADER_ROW = ['Timestamp', 'EmployeeName', 'LocationOrID', 'ModuleID', 'ModuleTitle', 'QuizScore', 'PassFail', 'Notes'];

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  ensureSheet();
  const template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Alterations Pinning Certification');
}

function ensureSheet() {
  const ss = getOrCreateSpreadsheet();
  let sheet = ss.getSheetByName(PROGRESS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRESS_SHEET_NAME);
    sheet.appendRow(HEADER_ROW);
  }
  const headers = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
  if (headers.join('') !== HEADER_ROW.join('')) {
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  }
  return sheet;
}

function getOrCreateSpreadsheet() {
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  return ss;
}

function saveModuleResult(payload) {
  if (!payload || !payload.employeeName || !payload.moduleId) {
    throw new Error('Invalid payload. Please provide employee name and module ID.');
  }
  const sheet = ensureSheet();
  const values = [
    new Date(),
    payload.employeeName.trim(),
    (payload.locationOrId || '').trim(),
    payload.moduleId,
    payload.moduleTitle || '',
    payload.quizScore || '',
    payload.passFail || '',
    payload.notes || ''
  ];
  sheet.appendRow(values);
  return {
    success: true,
    savedAt: values[0]
  };
}

function getModuleStatus(employeeName, locationOrId) {
  if (!employeeName) {
    return [];
  }
  const sheet = ensureSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(function(row) {
    const matchName = row[1] && row[1].toString().toLowerCase() === employeeName.toLowerCase();
    const matchLocation = !locationOrId || (row[2] && row[2].toString().toLowerCase() === locationOrId.toLowerCase());
    return matchName && matchLocation;
  });
  const latestByModule = {};
  rows.forEach(function(row) {
    const moduleId = row[3];
    const timestamp = row[0];
    if (!latestByModule[moduleId] || latestByModule[moduleId].timestamp < timestamp) {
      latestByModule[moduleId] = {
        timestamp: timestamp,
        employeeName: row[1],
        locationOrId: row[2],
        moduleId: moduleId,
        moduleTitle: row[4],
        quizScore: row[5],
        passFail: row[6],
        notes: row[7]
      };
    }
  });
  return Object.keys(latestByModule).map(function(key) { return latestByModule[key]; });
}

function getModulesCatalog() {
  return [
    {
      id: 'module-1',
      title: 'Customer Instruction & Measurement Philosophy',
      objectives: [
        'Use objective final measurements instead of subjective shorthand.',
        'Avoid +/- language; specify final measurements clearly.',
        'Guide customers to clarity and confirm understanding.',
        'Reinforce that hem is a garment part, not a measurement.'
      ]
    },
    {
      id: 'module-2',
      title: 'Pinning Tools & Safety',
      objectives: [
        'Choose correct pin types and default to safety pins for transport.',
        'Place pins horizontally to protect garments and staff.',
        'Apply pin safety rules to prevent accidents and garment damage.'
      ]
    },
    {
      id: 'module-3',
      title: 'Pinning by Garment Type',
      objectives: [
        'Smooth fabric before measuring and pinning.',
        'Balance inseams based on handedness questions for men.',
        'Verify customer-pinned garments with tape measure.',
        'Enforce one patch per garment and classify CSR vs tailor-only tasks.'
      ]
    },
    {
      id: 'module-4',
      title: 'SPOT POS Notes & Communication',
      objectives: [
        'Record clear, objective alteration notes in SPOT.',
        'Avoid vague language and confirm instructions with customers.',
        'Use annotated SPOT examples to avoid ambiguity.'
      ]
    },
    {
      id: 'module-5',
      title: 'Exceptions & Escalation',
      objectives: [
        'Identify garments CSRs must not pin.',
        'Escalate appropriately to tailors with respectful phrasing.',
        'Use visual cues and sorting practices to avoid mistakes.'
      ]
    }
  ];
}

function getLatestRecordsForPrint(employeeName, locationOrId) {
  const status = getModuleStatus(employeeName || '', locationOrId || '');
  return status;
}
