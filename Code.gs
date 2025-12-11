const SPREADSHEET_NAME = 'Alterations Pinning Certification';
const MODULE_RESULTS_SHEET = 'ModuleResults';
const MODULE_HEADERS = ['Timestamp', 'EmployeeName', 'LocationOrID', 'ModuleID', 'Score', 'Passed'];

// Reusable Google Doc certificate template for Dublin Cleaners
var CERTIFICATE_TEMPLATE_ID = '17yjalGF_nZEw_mWVQm9vlme_eoAYHLbBPw7nruiG1QQ';

// Folder to store generated certificates
var CERTIFICATE_FOLDER_NAME = 'Alterations Pinning Certificates';

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

function getOrCreateCertificateFolder_() {
  const folders = DriveApp.getFoldersByName(CERTIFICATE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(CERTIFICATE_FOLDER_NAME);
}

function generateCertificateFromTemplate(employeeName, employeeLocation) {
  if (!employeeName || !employeeName.trim()) {
    throw new Error('Employee name is required to generate a certificate.');
  }

  const cleanName = employeeName.trim();
  const cleanLocation = employeeLocation ? String(employeeLocation).trim() : '';
  const folder = getOrCreateCertificateFolder_();

  const status = getEmployeeCertificationStatus(cleanName);
  const isCertified = status && status.isCertified;

  const templateFile = DriveApp.getFileById(CERTIFICATE_TEMPLATE_ID);
  const today = new Date();
  const tz = Session.getScriptTimeZone();
  const dateStamp = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  const docName = 'Alterations Pinning Certificate - ' + cleanName + ' - ' + dateStamp;

  const newFile = templateFile.makeCopy(docName, folder);
  const newDoc = DocumentApp.openById(newFile.getId());

  const prettyDate = Utilities.formatDate(today, tz, 'MMMM d, yyyy');

  replacePlaceholderAcrossDoc_(newDoc, '{{EMPLOYEE_NAME}}', cleanName);
  replacePlaceholderAcrossDoc_(newDoc, '{{CERTIFICATE_DATE}}', prettyDate);
  replacePlaceholderAcrossDoc_(newDoc, '{{STORE_LOCATION}}', cleanLocation || '');
  replacePlaceholderAcrossDoc_(newDoc, '{{PROGRAM_NAME}}', 'Alterations Pinning Certification Program');

  newDoc.saveAndClose();

  const pdfBlob = newFile.getAs('application/pdf');
  pdfBlob.setName(docName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);

  return {
    docFileId: newFile.getId(),
    docFileUrl: newFile.getUrl(),
    pdfFileId: pdfFile.getId(),
    pdfFileUrl: pdfFile.getUrl(),
    isCertified: !!isCertified
  };
}

function replacePlaceholderAcrossDoc_(doc, placeholder, replacement) {
  const safeValue = replacement == null ? '' : replacement;
  const containers = [doc.getBody(), doc.getHeader(), doc.getFooter()];

  containers.forEach(function(container) {
    if (!container) return;
    let range = null;
    while (true) {
      range = container.findText(placeholder, range);
      if (!range) break;

      const element = range.getElement();
      if (!element || typeof element.editAsText !== 'function') continue;

      const text = element.asText();
      const start = range.getStartOffset();
      const end = range.getEndOffsetInclusive();
      const attrs = text.getAttributes(start) || {};

      text.deleteText(start, end);
      text.insertText(start, safeValue);

      if (safeValue.length > 0) {
        if (!attrs.foregroundColor) {
          attrs.foregroundColor = '#000000';
        }
        text.setAttributes(start, start + safeValue.length - 1, attrs);
      }
    }
  });
}

function buildCertificateContent_(body, employeeName, employeeLocationOrId, status, issuedOn) {
  body.clear();
  body.appendParagraph('Alterations Pinning Certification')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Certification of Completion')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Awarded to').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(employeeName)
    .setFontSize(18)
    .setBold(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const statement = `${employeeName} has completed the Dublin Cleaners Alterations Pinning Certification Program and is certified to pin garments for customers in-store.`;
  body.appendParagraph(statement)
    .setFontSize(12)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const table = body.appendTable([
    ['Issued On', Utilities.formatDate(issuedOn, Session.getScriptTimeZone() || 'America/New_York', 'MMMM d, yyyy')],
    ['Location / ID', employeeLocationOrId || '—'],
    ['Status', status.isCertified ? 'Certified' : 'In Progress']
  ]);
  table.getRow(0).editAsText().setBold(true);
  table.getRow(1).editAsText().setBold(true);
  table.getRow(2).editAsText().setBold(true);

  body.appendParagraph('Module Completion').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  const moduleList = body.appendListItem('');
  moduleList.clear();
  const allModules = getAllModules();
  const completed = status.completedModules || [];
  allModules.forEach(module => {
    const item = body.appendListItem(`${module.id} — ${module.title}`);
    item.setNestingLevel(0);
    item.setGlyphType(DocumentApp.GlyphType.BULLET);
    item.editAsText().setBold(completed.includes(module.id));
    if (completed.includes(module.id)) {
      item.appendText(' (Passed)');
    }
  });

  body.appendParagraph('Supervisor sign-off (if required): ________________________________')
    .setSpacingBefore(14)
    .setSpacingAfter(6)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function createCertificateFile(employeeName, employeeLocationOrId) {
  if (!employeeName) {
    throw new Error('Employee name is required to create a certificate.');
  }

  const status = getEmployeeCertificationStatus(employeeName);
  if (!status.isCertified) {
    throw new Error('Employee must complete all modules before creating a certificate.');
  }

  const folder = getOrCreateCertificateFolder_();
  const issuedOn = new Date();
  const tz = Session.getScriptTimeZone() || 'America/New_York';
  const dateStamp = Utilities.formatDate(issuedOn, tz, 'yyyy-MM-dd');
  const baseName = `Alterations Pinning Certificate - ${employeeName.trim()} - ${dateStamp}`;

  const doc = DocumentApp.create(baseName);
  buildCertificateContent_(doc.getBody(), employeeName.trim(), (employeeLocationOrId || '').trim(), status, issuedOn);
  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  const parents = docFile.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== folder.getId()) {
      parent.removeFile(docFile);
    }
  }

  const pdfBlob = docFile.getAs('application/pdf').setName(`${baseName}.pdf`);
  const pdfFile = folder.createFile(pdfBlob);

  return {
    fileId: pdfFile.getId(),
    fileUrl: pdfFile.getUrl(),
    docId: docFile.getId(),
    docUrl: docFile.getUrl(),
    folderUrl: folder.getUrl(),
    issuedOn: issuedOn,
    employeeName: employeeName.trim()
  };
}

function findLatestCertificateFile_(employeeName) {
  if (!employeeName) return null;

  const folder = getOrCreateCertificateFolder_();
  const files = folder.getFiles();
  const prefix = `Alterations Pinning Certificate - ${employeeName.trim()} -`;

  let latest = null;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().startsWith(prefix)) {
      const updated = file.getLastUpdated();
      if (!latest || (updated && updated > latest.updated)) {
        latest = {
          fileId: file.getId(),
          fileUrl: file.getUrl(),
          fileName: file.getName(),
          updated: updated || new Date(0)
        };
      }
    }
  }

  return latest;
}

function ensureCertificateFile(employeeName, employeeLocationOrId) {
  if (!employeeName) {
    throw new Error('Employee name is required to create a certificate.');
  }

  const status = getEmployeeCertificationStatus(employeeName);
  if (!status.isCertified) {
    throw new Error('Employee must complete all modules before creating a certificate.');
  }

  const existing = findLatestCertificateFile_(employeeName);
  if (existing) {
    return {
      fileId: existing.fileId,
      fileUrl: existing.fileUrl,
      fileName: existing.fileName,
      employeeName: employeeName.trim(),
      status: 'existing'
    };
  }

  const created = createCertificateFile(employeeName, employeeLocationOrId);
  return {
    fileId: created.fileId,
    fileUrl: created.fileUrl,
    fileName: `${created.employeeName} (new)`,
    employeeName: created.employeeName,
    status: 'created'
  };
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
