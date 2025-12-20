const SPREADSHEET_NAME = 'Alterations Pinning Certification';
const MODULE_RESULTS_SHEET = 'ModuleResults';
const QUIZ_ATTEMPTS_SHEET = 'QuizAttempts';
const EMPLOYEE_LOCKOUTS_SHEET = 'EmployeeLockouts';
const LOCKOUT_DURATION_MS = 72 * 60 * 60 * 1000;
const MODULE_HEADERS = ['Timestamp', 'EmployeeName', 'LocationOrID', 'ModuleID', 'Score', 'Passed'];
const QUIZ_ATTEMPT_HEADERS = [
  'Timestamp',
  'AttemptId',
  'EmployeeName',
  'LocationOrID',
  'ModuleID',
  'QuestionNumber',
  'QuestionId',
  'QuestionText',
  'SelectedOptionId',
  'SelectedOptionLabel',
  'CorrectOptionId',
  'CorrectOptionLabel',
  'IsCorrect',
  'CorrectCount',
  'TotalQuestions',
  'ScorePercent',
  'Passed'
];
const EMPLOYEE_LOCKOUT_HEADERS = [
  'EmployeeName',
  'EmployeeLocationOrId',
  'LockoutUntilIso',
  'LockoutUntilEpochMs',
  'Reason',
  'LastUpdatedTimestamp'
];

// Reusable Google Doc certificate template for Dublin Cleaners
var CERTIFICATE_TEMPLATE_ID = '17yjalGF_nZEw_mWVQm9vlme_eoAYHLbBPw7nruiG1QQ';

// Dublin Cleaners Certificate Assets
var CERTIFICATE_LOGO_URL = 'https://www.dublincleaners.com/wp-content/uploads/2025/06/LogosHQ.png';
var CERTIFICATE_BORDER_URL = 'https://www.dublincleaners.com/wp-content/uploads/2025/12/1Border.png';

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
  const spreadsheet = existing.hasNext()
    ? SpreadsheetApp.open(existing.next())
    : SpreadsheetApp.create(SPREADSHEET_NAME);
  getOrCreateModuleResultsSheet_(spreadsheet);
  getOrCreateQuizAttemptsSheet_(spreadsheet);
  getOrCreateEmployeeLockoutsSheet_(spreadsheet);
  return spreadsheet;
}

function getOrCreateSheet_(spreadsheet, sheetName, headers) {
  const ss = spreadsheet || getOrCreateCertificationSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (headers && headers.length) {
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    const existingHeaders = headerRange.getValues()[0];
    const headersMatch = existingHeaders.join('') === headers.join('');
    if (!headersMatch) {
      headerRange.setValues([headers]);
    }
  }
  return sheet;
}

function getOrCreateModuleResultsSheet_(spreadsheet) {
  return getOrCreateSheet_(spreadsheet, MODULE_RESULTS_SHEET, MODULE_HEADERS);
}

function getOrCreateQuizAttemptsSheet_(spreadsheet) {
  return getOrCreateSheet_(spreadsheet, QUIZ_ATTEMPTS_SHEET, QUIZ_ATTEMPT_HEADERS);
}

function getOrCreateEmployeeLockoutsSheet_(spreadsheet) {
  return getOrCreateSheet_(spreadsheet, EMPLOYEE_LOCKOUTS_SHEET, EMPLOYEE_LOCKOUT_HEADERS);
}

function appendModuleResultRow_(timestamp, moduleId, employeeName, employeeLocationOrId, score, passed, spreadsheet) {
  const sheet = getOrCreateModuleResultsSheet_(spreadsheet);
  const row = [
    timestamp,
    employeeName.trim(),
    (employeeLocationOrId || '').trim(),
    moduleId,
    typeof score === 'number' ? score : '',
    !!passed
  ];
  sheet.appendRow(row);
  return row;
}

function normalizeEmployeeValue_(value) {
  return (value || '').toString().trim().toLowerCase();
}

function findEmployeeLockoutRowIndex_(sheet, employeeName, employeeLocationOrId) {
  const data = sheet.getDataRange().getValues();
  const targetName = normalizeEmployeeValue_(employeeName);
  const targetLocation = normalizeEmployeeValue_(employeeLocationOrId);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (normalizeEmployeeValue_(row[0]) === targetName
      && normalizeEmployeeValue_(row[1]) === targetLocation) {
      return i + 1;
    }
  }
  return null;
}

function formatDurationMs_(durationMs) {
  const safeMs = Math.max(0, durationMs || 0);
  const totalSeconds = Math.floor(safeMs / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  return `${hours}:${('0' + minutes).slice(-2)}:${('0' + seconds).slice(-2)}`;
}

function setEmployeeLockout_(employeeName, employeeLocationOrId, reason, durationMs) {
  const ss = getOrCreateCertificationSpreadsheet_();
  const sheet = getOrCreateEmployeeLockoutsSheet_(ss);
  const lockoutUntil = new Date(Date.now() + (durationMs || LOCKOUT_DURATION_MS));
  const rowValues = [
    employeeName.trim(),
    (employeeLocationOrId || '').trim(),
    lockoutUntil.toISOString(),
    lockoutUntil.getTime(),
    reason || 'Failed quiz attempt',
    new Date()
  ];
  const existingIndex = findEmployeeLockoutRowIndex_(sheet, employeeName, employeeLocationOrId);
  if (existingIndex) {
    sheet.getRange(existingIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return rowValues;
}

function getEmployeeLockoutStatus(employeeName, employeeLocationOrId) {
  if (!employeeName) {
    return {
      isLocked: false,
      lockoutUntilEpochMs: null,
      lockoutUntilIso: null,
      msRemaining: 0,
      isCertified: false
    };
  }

  const normalizedName = employeeName.trim();
  const normalizedLocation = (employeeLocationOrId || '').trim();
  const status = getEmployeeCertificationStatus(normalizedName);
  if (status && status.isCertified) {
    return {
      isLocked: false,
      lockoutUntilEpochMs: null,
      lockoutUntilIso: null,
      msRemaining: 0,
      isCertified: true
    };
  }

  const sheet = getOrCreateEmployeeLockoutsSheet_();
  const data = sheet.getDataRange().getValues();
  const targetName = normalizeEmployeeValue_(normalizedName);
  const targetLocation = normalizeEmployeeValue_(normalizedLocation);

  let record = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (normalizeEmployeeValue_(row[0]) === targetName
      && normalizeEmployeeValue_(row[1]) === targetLocation) {
      record = row;
      break;
    }
  }

  if (!record) {
    return {
      isLocked: false,
      lockoutUntilEpochMs: null,
      lockoutUntilIso: null,
      msRemaining: 0,
      isCertified: false
    };
  }

  const lockoutUntilEpochMs = Number(record[3]) || null;
  const lockoutUntilIso = record[2] || null;
  const now = Date.now();
  const isLocked = lockoutUntilEpochMs && now < lockoutUntilEpochMs;
  return {
    isLocked: !!isLocked,
    lockoutUntilEpochMs: lockoutUntilEpochMs,
    lockoutUntilIso: lockoutUntilIso,
    msRemaining: isLocked ? Math.max(lockoutUntilEpochMs - now, 0) : 0,
    isCertified: false
  };
}

function saveModuleResult(moduleId, employeeName, employeeLocationOrId, score, passed) {
  if (!employeeName || !moduleId) {
    throw new Error('Employee name and module ID are required.');
  }

  const normalizedName = employeeName.trim();
  const normalizedLocation = (employeeLocationOrId || '').trim();
  const lockoutStatus = getEmployeeLockoutStatus(normalizedName, normalizedLocation);
  if (lockoutStatus.isLocked && !lockoutStatus.isCertified) {
    throw new Error(`Quiz lockout active. You can retry in ${formatDurationMs_(lockoutStatus.msRemaining)}.`);
  }

  const statusBefore = getEmployeeCertificationStatus(normalizedName);
  const timestamp = new Date();
  appendModuleResultRow_(timestamp, moduleId, normalizedName, normalizedLocation, score, passed);
  if (passed === false && statusBefore && !statusBefore.isCertified) {
    setEmployeeLockout_(normalizedName, normalizedLocation, 'Failed module quiz', LOCKOUT_DURATION_MS);
  }
  return { savedAt: timestamp };
}

function saveModuleAttempt(moduleId, employeeName, employeeLocationOrId, score, passed, correctCount, totalQuestions, answers) {
  if (!employeeName || !moduleId) {
    throw new Error('Employee name and module ID are required.');
  }

  const normalizedName = employeeName.trim();
  const normalizedLocation = (employeeLocationOrId || '').trim();
  const lockoutStatus = getEmployeeLockoutStatus(normalizedName, normalizedLocation);
  if (lockoutStatus.isLocked && !lockoutStatus.isCertified) {
    throw new Error(`Quiz lockout active. You can retry in ${formatDurationMs_(lockoutStatus.msRemaining)}.`);
  }

  const statusBefore = getEmployeeCertificationStatus(normalizedName);
  const ss = getOrCreateCertificationSpreadsheet_();
  const timestamp = new Date();
  const safeScore = typeof score === 'number' ? score : '';
  const safeCorrectCount = typeof correctCount === 'number' ? correctCount : 0;
  const safeTotalQuestions = typeof totalQuestions === 'number' && totalQuestions > 0 ? totalQuestions : 6;
  const attemptId = typeof Utilities !== 'undefined' && Utilities.getUuid ? Utilities.getUuid() : `${Date.now()}-${Math.floor(Math.random() * 100000)}`;

  appendModuleResultRow_(timestamp, moduleId, normalizedName, normalizedLocation, safeScore, passed, ss);

  const sheet = getOrCreateQuizAttemptsSheet_(ss);
  const responseRows = (answers || []).map(function(answer, idx) {
    return [
      timestamp,
      attemptId,
      normalizedName,
      normalizedLocation,
      moduleId,
      answer.questionNumber || idx + 1,
      answer.questionId || '',
      answer.questionText || '',
      answer.selectedOptionId || '',
      answer.selectedOptionLabel || '',
      answer.correctOptionId || '',
      answer.correctOptionLabel || '',
      answer.isCorrect === true,
      safeCorrectCount,
      safeTotalQuestions,
      safeScore,
      !!passed
    ];
  });

  const rowsToSave = responseRows.length ? responseRows : [[
    timestamp,
    attemptId,
    normalizedName,
    normalizedLocation,
    moduleId,
    '',
    '',
    'No question data captured',
    '',
    '',
    '',
    '',
    false,
    safeCorrectCount,
    safeTotalQuestions,
    safeScore,
    !!passed
  ]];

  sheet.getRange(sheet.getLastRow() + 1, 1, rowsToSave.length, QUIZ_ATTEMPT_HEADERS.length).setValues(rowsToSave);

  if (passed === false && statusBefore && !statusBefore.isCertified) {
    setEmployeeLockout_(normalizedName, normalizedLocation, 'Failed module quiz', LOCKOUT_DURATION_MS);
  }

  return { savedAt: timestamp, attemptId: attemptId };
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

function getCertificateBorderBlob_() {
  const borderBlob = UrlFetchApp.fetch(CERTIFICATE_BORDER_URL).getBlob();
  borderBlob.setName('border.png');
  return borderBlob;
}

function applyCertificateBands_(doc, borderBlob) {
  const safeBlob = borderBlob || getCertificateBorderBlob_();

  const header = doc.getHeader() || doc.addHeader();
  header.clear();
  const headerPara = header.appendParagraph('');
  headerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  headerPara.appendInlineImage(safeBlob).setWidth(600);

  const footer = doc.getFooter() || doc.addFooter();
  footer.clear();
  const footerPara = footer.appendParagraph('');
  footerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footerPara.appendInlineImage(safeBlob).setWidth(600);
}

function findExistingCertificateArtifacts_(employeeName) {
  const cleanName = employeeName ? employeeName.toString().trim() : '';
  if (!cleanName) return null;

  const folder = getOrCreateCertificateFolder_();
  const prefix = 'Alterations Pinning Certificate - ' + cleanName + ' -';
  const files = folder.getFiles();

  let latestPdf = null;
  let latestDoc = null;

  while (files.hasNext()) {
    const file = files.next();
    if (!file.getName().startsWith(prefix)) continue;

    const mime = file.getMimeType();
    const updated = typeof file.getLastUpdated === 'function' ? file.getLastUpdated() : null;
    const record = {
      file: file,
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: file.getName(),
      updated: updated || new Date(0)
    };

    if (mime === MimeType.PDF || mime === 'application/pdf') {
      if (!latestPdf || record.updated > latestPdf.updated) {
        latestPdf = record;
      }
    } else if (mime === MimeType.GOOGLE_DOCS || mime === 'application/vnd.google-apps.document') {
      if (!latestDoc || record.updated > latestDoc.updated) {
        latestDoc = record;
      }
    }
  }

  if (!latestPdf && !latestDoc) return null;
  return { pdf: latestPdf, doc: latestDoc };
}

function generateCertificatePDF(employeeName, employeeLocationOrId) {
  const cleanName = employeeName ? employeeName.toString().trim() : '';
  if (!cleanName) {
    throw new Error('Employee name is required to generate a certificate.');
  }

  const cleanLocation = employeeLocationOrId ? employeeLocationOrId.toString().trim() : '';
  let status = null;
  try {
    status = getEmployeeCertificationStatus(cleanName);
  } catch (err) {
    // If status lookup fails, proceed without blocking certificate generation.
    status = null;
  }

  const existing = findExistingCertificateArtifacts_(cleanName);
  if (existing && existing.pdf) {
    return {
      docFileId: existing.doc ? existing.doc.fileId : null,
      docFileUrl: existing.doc ? existing.doc.fileUrl : null,
      pdfFileId: existing.pdf.fileId,
      pdfFileUrl: existing.pdf.fileUrl,
      isCertified: !!(status && status.isCertified)
    };
  }

  const folder = getOrCreateCertificateFolder_();
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayIso = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  const prettyDate = Utilities.formatDate(today, tz, 'MMMM d, yyyy');
  const docName = 'Alterations Pinning Certificate - ' + cleanName + ' - ' + todayIso;

  const doc = DocumentApp.create(docName);
  const docId = doc.getId();
  const file = DriveApp.getFileById(docId);
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.clear();

  const borderBlob = getCertificateBorderBlob_();
  applyCertificateBands_(doc, borderBlob);

  const table = body.appendTable([['']]);
  table.setBorderWidth(0);
  const cell = table.getCell(0, 0);
  cell.setPaddingTop(20);
  cell.setPaddingBottom(20);
  cell.setPaddingLeft(20);
  cell.setPaddingRight(20);

  const title1 = cell.appendParagraph('Dublin Cleaners');
  title1.setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const title2 = cell.appendParagraph('Alterations Pinning Certification');
  title2.setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const subtitle = cell.appendParagraph('Official Certification');
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  cell.appendParagraph('');

  const label = cell.appendParagraph('This certifies that');
  label.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const namePara = cell.appendParagraph(cleanName);
  namePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setBold(true)
    .setFontSize(18);

  const bodyText = 'has successfully completed the Dublin Cleaners Alterations Pinning Certification Program and is hereby certified and authorized to pin garments for our customers in-store in accordance with our quality, safety, and service standards.';
  const bodyPara = cell.appendParagraph(bodyText);
  bodyPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setFontSize(11);

  const datePara = cell.appendParagraph('Date: ' + prettyDate);
  datePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setFontSize(10);

  if (cleanLocation) {
    const locPara = cell.appendParagraph('Store / Location: ' + cleanLocation);
    locPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontSize(10);
  }

  cell.appendParagraph('');

  const logoBlob = UrlFetchApp.fetch(CERTIFICATE_LOGO_URL).getBlob();
  logoBlob.setName('logo.png');
  const logoPara = cell.appendParagraph('');
  const logoImage = logoPara.appendInlineImage(logoBlob);
  logoImage.setWidth(150).setHeight(150);
  logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(6)
    .setSpacingAfter(0);

  doc.saveAndClose();

  const pdfBlob = file.getAs('application/pdf');
  pdfBlob.setName(docName + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);

  return {
    docFileId: file.getId(),
    docFileUrl: file.getUrl(),
    pdfFileId: pdfFile.getId(),
    pdfFileUrl: pdfFile.getUrl(),
    isCertified: !!(status && status.isCertified)
  };
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

  applyCertificateBands_(newDoc, getCertificateBorderBlob_());

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
        attrs.foregroundColor = '#000000';
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

  const statement = `${employeeName} has completed the Official Dublin Cleaners Alterations Pinning Certification Program and is certified to pin garments.`;
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
    item.appendText(completed.includes(module.id) ? ' (Passed)' : ' (Pending)');
  });

  body.appendParagraph('Supervisor sign-off (required): ________________________________')
    .setSpacingBefore(18)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Manager signature: ________________________________')
    .setSpacingBefore(12)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Manager name / comments: ________________________________________________')
    .setSpacingBefore(8)
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
