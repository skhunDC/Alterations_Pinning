const SPREADSHEET_NAME = 'Alterations Pinning Certification';
const MODULE_RESULTS_SHEET = 'ModuleResults';
const MODULE_HEADERS = ['Timestamp', 'EmployeeName', 'LocationOrID', 'ModuleID', 'Score', 'Passed'];

function include(filename) {
  const name = (filename || '').toString().trim();
  if (!name) {
    Logger.log('Include called without a filename; returning empty string.');
    return '';
  }
  try {
    return HtmlService.createHtmlOutputFromFile(name).getContent();
  } catch (err) {
    Logger.log(`Include failed for "${name}": ${err.message}`);
    return '';
  }
}

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setTitle('Alterations Pinning Certification Program')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateCertificationSpreadsheet_() {
  const existing = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (existing.hasNext()) {
    return SpreadsheetApp.open(existing.next());
  }
  return SpreadsheetApp.create(SPREADSHEET_NAME);
}

function getOrCreateModuleResultsSheet_() {
  const ss = getOrCreateCertificationSpreadsheet_();
  let sheet = ss.getSheetByName(MODULE_RESULTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MODULE_RESULTS_SHEET);
  }
  const headerRange = sheet.getRange(1, 1, 1, MODULE_HEADERS.length);
  const headers = headerRange.getValues()[0];
  const headersMatch = headers.join('') === MODULE_HEADERS.join('');
  if (!headersMatch) {
    headerRange.setValues([MODULE_HEADERS]);
  }
  return sheet;
}

function saveModuleResult(moduleId, employeeName, employeeLocationOrId, score, passed) {
  if (!employeeName || !moduleId) {
    throw new Error('Employee name and module ID are required.');
  }
  const sheet = getOrCreateModuleResultsSheet_();
  const row = [
    new Date(),
    employeeName.trim(),
    (employeeLocationOrId || '').trim(),
    moduleId,
    typeof score === 'number' ? score : '',
    !!passed
  ];
  sheet.appendRow(row);
  return { savedAt: row[0] };
}

function getEmployeeCertificationStatus(employeeName) {
  if (!employeeName) {
    return { completedModules: [], missingModules: ['M1', 'M2', 'M3', 'M4', 'M5'], isCertified: false };
  }
  const sheet = getOrCreateModuleResultsSheet_();
  const data = sheet.getDataRange().getValues();
  const records = data.slice(1).filter(function(row) {
    return row[1] && row[1].toString().trim().toLowerCase() === employeeName.trim().toLowerCase();
  });
  const latest = {};
  records.forEach(function(row) {
    const moduleId = row[3];
    const timestamp = row[0];
    if (!latest[moduleId] || latest[moduleId].timestamp < timestamp) {
      latest[moduleId] = {
        moduleId: moduleId,
        score: row[4],
        passed: row[5] === true || row[5] === 'TRUE',
        timestamp: timestamp
      };
    }
  });
  const allModules = ['M1', 'M2', 'M3', 'M4', 'M5'];
  const completedModules = allModules.filter(function(id) {
    return latest[id] && latest[id].passed === true;
  });
  const missingModules = allModules.filter(function(id) { return completedModules.indexOf(id) === -1; });
  return { completedModules: completedModules, missingModules: missingModules, isCertified: missingModules.length === 0 };
}

function getAllModules() {
  return [
    { id: 'M1', title: 'Customer Instruction & Measurement Philosophy' },
    { id: 'M2', title: 'Pinning Tools & Safety' },
    { id: 'M3', title: 'Pinning by Garment Type' },
    { id: 'M4', title: 'SPOT POS Notes & Communication' },
    { id: 'M5', title: 'Exceptions & Escalation' }
  ];
}

function getAllModuleResults() {
  const sheet = getOrCreateModuleResultsSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  return rows.map(row => ({
    timestamp: row[0],
    employeeName: row[1],
    employeeLocationOrId: row[2],
    moduleId: row[3],
    score: row[4],
    passed: row[5] === true || row[5] === 'TRUE'
  }));
}

function getSummaryByEmployee() {
  const results = getAllModuleResults();
  const summaryMap = {};
  results.forEach(entry => {
    const key = (entry.employeeName || '').trim().toLowerCase();
    if (!key) return;
    if (!summaryMap[key]) {
      summaryMap[key] = {
        employeeName: entry.employeeName,
        employeeLocationOrId: entry.employeeLocationOrId,
        modulesPassed: [],
        modulesFailed: [],
        lastUpdated: entry.timestamp
      };
    }
    const existing = summaryMap[key];
    if (entry.employeeLocationOrId && !existing.employeeLocationOrId) {
      existing.employeeLocationOrId = entry.employeeLocationOrId;
    }
    if (!existing.lastUpdated || (entry.timestamp && entry.timestamp > existing.lastUpdated)) {
      existing.lastUpdated = entry.timestamp;
    }
    if (entry.passed) {
      if (!existing.modulesPassed.includes(entry.moduleId)) {
        existing.modulesPassed.push(entry.moduleId);
      }
    } else {
      if (!existing.modulesFailed.includes(entry.moduleId)) {
        existing.modulesFailed.push(entry.moduleId);
      }
    }
  });

  return Object.values(summaryMap).map(item => {
    const allModules = ['M1', 'M2', 'M3', 'M4', 'M5'];
    const isCertified = allModules.every(id => item.modulesPassed.includes(id));
    return {
      employeeName: item.employeeName,
      employeeLocationOrId: item.employeeLocationOrId,
      modulesPassed: item.modulesPassed.sort(),
      modulesFailed: item.modulesFailed.sort(),
      isCertified: isCertified,
      lastUpdated: item.lastUpdated
    };
  });
}
