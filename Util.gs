function getOrCreateTrackerSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  const existingId = getScriptProperty_(PROPERTY_KEYS.TRACKER_SPREADSHEET_ID);
  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (err) {
      clearScriptProperty_(PROPERTY_KEYS.TRACKER_SPREADSHEET_ID);
    }
  }

  return SpreadsheetApp.create(getConfigString_('TRACKER_FILENAME'));
}

function getTrackerSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  const id = getScriptProperty_(PROPERTY_KEYS.TRACKER_SPREADSHEET_ID);
  if (!id) {
    throw new Error('Tracker spreadsheet is not configured. Run setupTracker() first.');
  }
  return SpreadsheetApp.openById(id);
}

function getTrackerSheet_() {
  const trackerSheetName = getConfigString_('TRACKER_SHEET_NAME');
  const sheet = getTrackerSpreadsheet_().getSheetByName(trackerSheetName);
  if (!sheet) {
    throw new Error('Missing tracker sheet "' + trackerSheetName + '". Run setupTracker() first.');
  }
  return sheet;
}

function getLogsSheet_() {
  const ss = getTrackerSpreadsheet_();
  const sheetName = getConfigString_('LOGS_SHEET_NAME');
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getClassificationLogSheet_() {
  const ss = getTrackerSpreadsheet_();
  const sheetName = getConfigString_('CLASSIFICATION_LOG_SHEET_NAME');
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
    sheet.getRange(1, 1, 1, CLASSIFICATION_LOG_HEADERS.length).setValues([CLASSIFICATION_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CLASSIFICATION_LOG_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.getRange(1, 1, sheet.getMaxRows(), CLASSIFICATION_LOG_HEADERS.length).createFilter();
  }
  return sheet;
}

function getDriveClassificationRegistrySheet_() {
  const ss = getTrackerSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.DRIVE_CLASSIFICATION_REGISTRY_SHEET_NAME)
    || ss.insertSheet(CONFIG.DRIVE_CLASSIFICATION_REGISTRY_SHEET_NAME);
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
    sheet.getRange(1, 1, 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).setValues([DRIVE_CLASSIFICATION_REGISTRY_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.getRange(1, 1, sheet.getMaxRows(), DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).createFilter();
  }
  applyDriveClassificationRegistryFormats_(sheet);
  return sheet;
}

function applyDriveClassificationRegistryFormats_(sheet) {
  if (sheet.getMaxRows() < 2) {
    return;
  }
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2, 2, sheet.getMaxRows() - 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length - 1).setNumberFormat('@');
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getMasterConfigSheet_(ss) {
  return ss.getSheetByName(CONFIG.MASTER_CONFIG_SHEET_NAME);
}

function clearMasterConfigCache_() {
  MASTER_CONFIG_CACHE_ = null;
}

function getRuntimeConfigSnapshot_() {
  const snapshot = {};
  MASTER_CONFIG_DEFINITIONS.forEach(function (definition) {
    snapshot[definition.key] = getConfigValue_(definition.key);
  });
  return snapshot;
}

function getConfigValue_(key) {
  const definition = getMasterConfigDefinition_(key);
  const valueType = definition ? definition.valueType : inferConfigValueType_(CONFIG[key]);
  const allowedValues = definition ? definition.allowedValues : '';
  const staticDefault = definition ? definition.defaultValue : CONFIG[key];
  let row = null;

  try {
    row = getMasterConfigRow_(key);
  } catch (err) {
    return staticDefault;
  }

  const candidates = [];
  if (row && !isConfigBlank_(row.value)) {
    candidates.push({ source: 'value', value: row.value });
  } else if (row && !isConfigBlank_(staticDefault)) {
    logInvalidConfigValue_(key, 'value', row.value, 'Blank value. Falling back to default.');
  }
  if (row && !isConfigBlank_(row.defaultValue)) {
    candidates.push({ source: 'default_value', value: row.defaultValue });
  }
  candidates.push({ source: 'static_default', value: staticDefault });

  for (let i = 0; i < candidates.length; i++) {
    const parsed = parseConfigValue_(candidates[i].value, valueType, allowedValues);
    if (parsed.valid) {
      return parsed.value;
    }
    if (candidates[i].source !== 'static_default') {
      logInvalidConfigValue_(key, candidates[i].source, candidates[i].value, parsed.reason);
    }
  }

  return staticDefault;
}

function getConfigString_(key) {
  const value = getConfigValue_(key);
  return value === null || value === undefined ? '' : String(value);
}

function getConfigInteger_(key) {
  const value = getConfigValue_(key);
  const number = Number(value);
  return isNaN(number) ? 0 : Math.floor(number);
}

function isFeatureEnabled_(key) {
  return getConfigValue_(key) === true;
}

function getMasterConfigDefinition_(key) {
  for (let i = 0; i < MASTER_CONFIG_DEFINITIONS.length; i++) {
    if (MASTER_CONFIG_DEFINITIONS[i].key === key) {
      return MASTER_CONFIG_DEFINITIONS[i];
    }
  }
  return null;
}

function getMasterConfigRow_(key) {
  const rows = getMasterConfigRows_();
  return rows[key] || null;
}

function getMasterConfigRows_() {
  if (MASTER_CONFIG_CACHE_) {
    return MASTER_CONFIG_CACHE_;
  }

  const rowsByKey = {};
  const ss = getTrackerSpreadsheet_();
  const sheet = getMasterConfigSheet_(ss);
  if (!sheet || sheet.getLastRow() < 2) {
    MASTER_CONFIG_CACHE_ = rowsByKey;
    return rowsByKey;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (value) {
    return String(value || '').trim();
  });
  const keyIndex = headers.indexOf('setting_key');
  const valueIndex = headers.indexOf('value');
  const defaultIndex = headers.indexOf('default_value');
  if (keyIndex < 0 || valueIndex < 0 || defaultIndex < 0) {
    MASTER_CONFIG_CACHE_ = rowsByKey;
    return rowsByKey;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  values.forEach(function (row) {
    const key = String(row[keyIndex] || '').trim();
    if (!key) {
      return;
    }
    rowsByKey[key] = {
      value: row[valueIndex],
      defaultValue: row[defaultIndex]
    };
  });

  MASTER_CONFIG_CACHE_ = rowsByKey;
  return rowsByKey;
}

function parseConfigValue_(rawValue, valueType, allowedValues) {
  if (valueType === 'boolean') {
    return parseBooleanConfigValue_(rawValue);
  }

  if (valueType === 'integer') {
    const number = Number(rawValue);
    if (!isFinite(number) || Math.floor(number) !== number || number < 0) {
      return { valid: false, reason: 'Expected a non-negative integer.' };
    }
    return { valid: true, value: number };
  }

  if (valueType === 'enum') {
    const choices = splitAllowedValues_(allowedValues);
    const value = String(rawValue || '').trim().toLowerCase();
    if (choices.indexOf(value) < 0) {
      return { valid: false, reason: 'Expected one of: ' + choices.join(', ') + '.' };
    }
    return { valid: true, value: value };
  }

  if (isConfigBlank_(rawValue)) {
    return { valid: false, reason: 'Expected a non-blank value.' };
  }
  return { valid: true, value: String(rawValue).trim() };
}

function parseBooleanConfigValue_(rawValue) {
  if (rawValue === true || rawValue === false) {
    return { valid: true, value: rawValue };
  }

  const value = String(rawValue || '').trim().toLowerCase();
  if (['true', 'yes', 'y', '1', 'enabled', 'on'].indexOf(value) >= 0) {
    return { valid: true, value: true };
  }
  if (['false', 'no', 'n', '0', 'disabled', 'off'].indexOf(value) >= 0) {
    return { valid: true, value: false };
  }
  return { valid: false, reason: 'Expected TRUE or FALSE.' };
}

function serializeConfigValue_(value, valueType) {
  if (valueType === 'boolean') {
    return value === true;
  }
  if (value === null || value === undefined) {
    return '';
  }
  return value;
}

function inferConfigValueType_(value) {
  if (typeof value === 'boolean') {
    return 'boolean';
  }
  if (typeof value === 'number') {
    return 'integer';
  }
  return 'string';
}

function isConfigBlank_(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function splitAllowedValues_(allowedValues) {
  return String(allowedValues || '')
    .split(',')
    .map(function (value) {
      return value.trim().toLowerCase();
    })
    .filter(Boolean);
}

function logInvalidConfigValue_(key, source, value, reason) {
  const warningKey = key + '|' + source + '|' + String(value);
  if (CONFIG_WARNING_CACHE_[warningKey]) {
    return;
  }
  CONFIG_WARNING_CACHE_[warningKey] = true;

  try {
    const ss = getTrackerSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.LOGS_SHEET_NAME) || ss.insertSheet(CONFIG.LOGS_SHEET_NAME);
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
      sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      'invalid_master_config_value',
      '',
      '',
      'Invalid Master Config value ignored for ' + key,
      safeJsonStringify_({
        settingKey: key,
        source: source,
        value: value,
        reason: reason
      })
    ]);
  } catch (err) {
    // Logging configuration errors must never block runtime fallback behavior.
  }
}

function getCachedDriveClassification_(file, record, docs) {
  const fileId = getDriveFileIdSafe_(file);
  if (!fileId || !record) {
    return null;
  }

  const docsSignature = buildRequiredDocsSignature_(docs);
  const key = buildDriveClassificationRegistryKey_(fileId, record.company_name, record.month, docsSignature);
  const cached = getDriveClassificationRegistryMap_()[key];
  if (!cached) {
    return null;
  }

  const currentFingerprint = buildDriveClassificationFileFingerprint_(file);
  return cached.file_fingerprint === currentFingerprint ? cached : null;
}

function clearDriveClassificationRegistryCache_() {
  DRIVE_CLASSIFICATION_REGISTRY_CACHE_ = null;
}

function rememberDriveClassification_(file, record, docs, source, classification) {
  if (!shouldRememberDriveClassification_(classification)) {
    return;
  }

  const fileId = getDriveFileIdSafe_(file);
  if (!fileId || !record) {
    return;
  }

  const docsSignature = buildRequiredDocsSignature_(docs);
  const fingerprint = buildDriveClassificationFileFingerprint_(file);
  const matchedDocKeys = (classification.matchedDocs || []).map(function (doc) {
    return doc.doc_key;
  }).join(',');
  const status = matchedDocKeys ? 'matched' : 'no_match';
  const fileUrl = getDriveFileUrlSafe_(file);
  const row = [
    new Date(),
    fileId,
    record.company_name || '',
    normalizeRegistryMonthValue_(record.month),
    docsSignature,
    fingerprint,
    source || '',
    status,
    matchedDocKeys,
    file.getName(),
    fileUrl,
    classification.reason || ''
  ];

  const sheet = getDriveClassificationRegistrySheet_();
  const key = buildDriveClassificationRegistryKey_(fileId, record.company_name, record.month, docsSignature);
  const registry = getDriveClassificationRegistryMap_();
  const existing = registry[key];
  const rowNumber = existing && existing.rowNumber ? existing.rowNumber : sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).setValues([row]);
  registry[key] = buildDriveClassificationRegistryRecordFromRow_(row, rowNumber);
}

function shouldRememberDriveClassification_(classification) {
  if (!classification) {
    return false;
  }
  if (classification.matchedDocs && classification.matchedDocs.length > 0) {
    return true;
  }
  if (classification.raw) {
    return true;
  }

  const reason = String(classification.reason || '').toLowerCase();
  if (reason.indexOf('no required document types accept extension') === 0) {
    return true;
  }
  return false;
}

function getDocsFromCachedDriveClassification_(cached, docs) {
  const docsByKey = {};
  docs.forEach(function (doc) {
    docsByKey[doc.doc_key] = doc;
  });

  return String(cached.matched_doc_keys || '')
    .split(',')
    .map(function (key) {
      return docsByKey[String(key || '').trim()];
    })
    .filter(Boolean);
}

function getDriveClassificationRegistryMap_() {
  if (DRIVE_CLASSIFICATION_REGISTRY_CACHE_) {
    return DRIVE_CLASSIFICATION_REGISTRY_CACHE_;
  }

  const registry = {};
  const sheet = getDriveClassificationRegistrySheet_();
  if (sheet.getLastRow() < 2) {
    DRIVE_CLASSIFICATION_REGISTRY_CACHE_ = registry;
    return registry;
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).getValues();
  rows.forEach(function (row, index) {
    const record = buildDriveClassificationRegistryRecordFromRow_(row, index + 2);
    if (!record.file_id || !record.docs_signature) {
      return;
    }
    const key = buildDriveClassificationRegistryKey_(record.file_id, record.company_name, record.month, record.docs_signature);
    registry[key] = choosePreferredDriveClassificationRegistryRecord_(registry[key], record);
  });

  DRIVE_CLASSIFICATION_REGISTRY_CACHE_ = registry;
  return registry;
}

function dedupeDriveClassificationRegistry_() {
  const sheet = getDriveClassificationRegistrySheet_();
  if (sheet.getLastRow() < 3) {
    return;
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).getValues();
  const recordsByKey = {};
  const uniqueKeys = [];
  rows.forEach(function (row, index) {
    const record = buildDriveClassificationRegistryRecordFromRow_(row, index + 2);
    if (!record.file_id || !record.docs_signature) {
      uniqueKeys.push('row-' + index);
      recordsByKey['row-' + index] = record;
      return;
    }

    const key = buildDriveClassificationRegistryKey_(record.file_id, record.company_name, record.month, record.docs_signature);
    if (!recordsByKey[key]) {
      uniqueKeys.push(key);
      recordsByKey[key] = record;
      return;
    }
    recordsByKey[key] = choosePreferredDriveClassificationRegistryRecord_(recordsByKey[key], record);
  });

  const dedupedRows = uniqueKeys.map(function (key) {
    return DRIVE_CLASSIFICATION_REGISTRY_HEADERS.map(function (header) {
      return recordsByKey[key][header] || '';
    });
  });

  if (dedupedRows.length === rows.length) {
    return;
  }

  sheet.getRange(2, 1, sheet.getLastRow() - 1, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).clearContent();
  sheet.getRange(2, 1, dedupedRows.length, DRIVE_CLASSIFICATION_REGISTRY_HEADERS.length).setValues(dedupedRows);
  DRIVE_CLASSIFICATION_REGISTRY_CACHE_ = null;
  logEvent_('drive_classification_registry_deduped', '', '', 'Removed duplicate Drive classification registry rows', {
    beforeRows: rows.length,
    afterRows: dedupedRows.length,
    removedRows: rows.length - dedupedRows.length
  });
}

function choosePreferredDriveClassificationRegistryRecord_(current, candidate) {
  if (!current) {
    return candidate;
  }
  if (!candidate) {
    return current;
  }

  const currentHasFingerprint = !isConfigBlank_(current.file_fingerprint);
  const candidateHasFingerprint = !isConfigBlank_(candidate.file_fingerprint);
  if (candidateHasFingerprint && !currentHasFingerprint) {
    return candidate;
  }
  if (currentHasFingerprint && !candidateHasFingerprint) {
    return current;
  }

  return getRegistryTimestampValue_(candidate) >= getRegistryTimestampValue_(current)
    ? candidate
    : current;
}

function getRegistryTimestampValue_(record) {
  const timestamp = record ? record.timestamp : null;
  if (Object.prototype.toString.call(timestamp) === '[object Date]') {
    return timestamp.getTime();
  }
  const parsed = new Date(timestamp);
  const time = parsed.getTime();
  if (!isNaN(time)) {
    return time;
  }
  return Number(record && record.rowNumber ? record.rowNumber : 0);
}

function buildDriveClassificationRegistryRecordFromRow_(row, rowNumber) {
  const record = {};
  DRIVE_CLASSIFICATION_REGISTRY_HEADERS.forEach(function (header, index) {
    record[header] = row[index];
  });
  record.rowNumber = rowNumber || 0;
  return record;
}

function buildDriveClassificationRegistryKey_(fileId, companyName, month, docsSignature) {
  return [
    String(fileId || '').trim(),
    normalizeText_(companyName),
    normalizeRegistryMonthValue_(month),
    String(docsSignature || '').trim()
  ].join('|');
}

function normalizeRegistryMonthValue_(month) {
  if (Object.prototype.toString.call(month) === '[object Date]' && !isNaN(month.getTime())) {
    return formatMonth_(month);
  }
  return String(month || '').trim();
}

function buildRequiredDocsSignature_(docs) {
  return (docs || []).map(function (doc) {
    return [
      doc.doc_key,
      normalizeText_(doc.display_name),
      normalizeText_(doc.accepted_extensions)
    ].join(':');
  }).sort().join('|');
}

function buildDriveClassificationFileFingerprint_(file) {
  return [
    getDriveFileIdSafe_(file),
    getDriveFileMimeTypeSafe_(file),
    getDriveFileSizeSafe_(file),
    getDriveFileLastUpdatedSafe_(file)
  ].join('|');
}

function getDriveFileIdSafe_(file) {
  try {
    return file.getId();
  } catch (err) {
    return '';
  }
}

function getDriveFileMimeTypeSafe_(file) {
  try {
    return file.getMimeType();
  } catch (err) {
    return '';
  }
}

function getDriveFileSizeSafe_(file) {
  try {
    return typeof file.getSize === 'function' ? file.getSize() : '';
  } catch (err) {
    return '';
  }
}

function getDriveFileLastUpdatedSafe_(file) {
  try {
    const updated = file.getLastUpdated();
    return updated ? updated.getTime() : '';
  } catch (err) {
    return '';
  }
}

function getDriveFileUrlSafe_(file) {
  try {
    return file.getUrl();
  } catch (err) {
    return '';
  }
}

function getRequiredDocs_() {
  const ss = getTrackerSpreadsheet_();
  const sheet = ss.getSheetByName(getConfigString_('REQUIRED_DOCS_SHEET_NAME'));
  if (!sheet || sheet.getLastRow() < 2) {
    return CONFIG.DEFAULT_DOCS.slice();
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const docs = values
    .filter(function (row) {
      return row[0] && row[1];
    })
    .map(function (row) {
      return {
        doc_key: sanitizeDocKey_(row[0]),
        display_name: String(row[1]).trim(),
        accepted_extensions: String(row[2] || '').trim()
      };
    });

  return docs.length ? docs : CONFIG.DEFAULT_DOCS.slice();
}

function buildSubmissionHeaders_(docs) {
  const docHeaders = [];
  docs.forEach(function (doc) {
    docHeaders.push('doc_' + doc.doc_key + '_status');
    docHeaders.push('doc_' + doc.doc_key + '_date');
  });
  return BASE_SUBMISSION_HEADERS.concat(docHeaders).concat(TRAILING_SUBMISSION_HEADERS);
}

function ensureTrackerColumns_(sheet, docs) {
  const desiredHeaders = buildSubmissionHeaders_(docs);
  const currentHeaders = getHeaders_(sheet);

  if (currentHeaders.length === 0 || currentHeaders[0] === '') {
    sheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    return;
  }

  desiredHeaders.forEach(function (header) {
    if (currentHeaders.indexOf(header) === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(header);
    }
  });

  backfillMissingDeadlines_(sheet, getHeaders_(sheet));
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) {
    return [];
  }
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (value) {
    return String(value || '').trim();
  });
}

function getHeaderMap_(sheet) {
  const headers = getHeaders_(sheet);
  const map = {};
  headers.forEach(function (header, index) {
    if (header) {
      map[header] = index + 1;
    }
  });
  return map;
}

function setCellByHeader_(sheet, rowNumber, header, value) {
  const map = getHeaderMap_(sheet);
  if (!map[header]) {
    throw new Error('Missing tracker column: ' + header);
  }
  sheet.getRange(rowNumber, map[header]).setValue(value);
}

function getCellByHeader_(sheet, rowNumber, header) {
  const map = getHeaderMap_(sheet);
  if (!map[header]) {
    return '';
  }
  return sheet.getRange(rowNumber, map[header]).getValue();
}

function applySubmissionFormulas_(sheet, rowNumber, docs) {
  const headers = getHeaders_(sheet);
  setCellByHeader_(sheet, rowNumber, 'all_complete', buildAllCompleteFormula_(rowNumber, docs, headers));
  setCellByHeader_(sheet, rowNumber, 'days_overdue', buildDaysOverdueFormula_(rowNumber, headers));
}

function buildAllCompleteFormula_(rowNumber, docs, headers) {
  const tests = docs.map(function (doc) {
    const col = headers.indexOf('doc_' + doc.doc_key + '_status') + 1;
    return columnToLetter_(col) + rowNumber + '="received"';
  });
  return '=AND(' + tests.join(',') + ')';
}

function buildDaysOverdueFormula_(rowNumber, headers) {
  const allCompleteCol = columnToLetter_(headers.indexOf('all_complete') + 1);
  const deadlineCol = columnToLetter_(headers.indexOf('deadline') + 1);
  return '=IF(OR(' + allCompleteCol + rowNumber + ',' + deadlineCol + rowNumber + '=""),0,MAX(0,TODAY()-' + deadlineCol + rowNumber + '))';
}

function applyStatusValidation_(sheet, docs) {
  const headers = getHeaders_(sheet);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['missing', 'received'], true)
    .setAllowInvalid(false)
    .build();

  docs.forEach(function (doc) {
    const col = headers.indexOf('doc_' + doc.doc_key + '_status') + 1;
    if (col > 0) {
      sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setDataValidation(validation);
    }
  });
}

function applyDeadlineValidation_(sheet, headers) {
  const deadlineCol = headers.indexOf('deadline') + 1;
  if (deadlineCol <= 0) {
    return;
  }

  const validation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .setHelpText('Use a real date. Editing this deadline resets reminders and escalation timing for incomplete submissions.')
    .build();
  sheet.getRange(2, deadlineCol, sheet.getMaxRows() - 1, 1)
    .setDataValidation(validation)
    .setNumberFormat('yyyy-mm-dd');
  sheet.getRange(1, deadlineCol)
    .setNote('Editable submission deadline. Defaults from the reporting month; changing it recalculates days overdue and resets reminder/escalation timing for incomplete rows.')
    .setBackground('#fff2cc');
}

function applyTrackerNumberFormats_(sheet, headers) {
  const monthCol = headers.indexOf('month') + 1;
  const deadlineCol = headers.indexOf('deadline') + 1;
  const daysCol = headers.indexOf('days_overdue') + 1;
  const reminderCountCol = headers.indexOf('reminder_count') + 1;

  if (monthCol > 0) {
    sheet.getRange(2, monthCol, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  if (deadlineCol > 0) {
    sheet.getRange(2, deadlineCol, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  ['last_reminder_at', 'escalated_at'].forEach(function (header) {
    const col = headers.indexOf(header) + 1;
    if (col > 0) {
      sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }
  });
  headers.forEach(function (header, index) {
    if (header.indexOf('doc_') === 0 && header.slice(-5) === '_date') {
      sheet.getRange(2, index + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }
  });
  if (daysCol > 0) {
    sheet.getRange(2, daysCol, sheet.getMaxRows() - 1, 1).setNumberFormat('0');
  }
  if (reminderCountCol > 0) {
    sheet.getRange(2, reminderCountCol, sheet.getMaxRows() - 1, 1).setNumberFormat('0');
  }
}

function handleMasterConfigEdit_(e) {
  if (!e || !e.range) {
    return;
  }
  if (e.range.getSheet().getName() === CONFIG.MASTER_CONFIG_SHEET_NAME) {
    clearMasterConfigCache_();
  }
}

function handleSubmissionTrackerEdit_(e) {
  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== getConfigString_('TRACKER_SHEET_NAME') || range.getRow() < 2) {
    return;
  }

  const headers = getHeaders_(sheet);
  const deadlineCol = headers.indexOf('deadline') + 1;
  if (deadlineCol <= 0 || deadlineCol < range.getColumn() || deadlineCol > range.getLastColumn()) {
    return;
  }

  const docs = getRequiredDocs_();
  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getLastRow();
  for (let row = firstRow; row <= lastRow; row++) {
    applySubmissionFormulas_(sheet, row, docs);
    const record = getRecordAtRow_(sheet, row);
    if (!isRecordComplete_(record, docs)) {
      resetReminderScheduleForRow_(sheet, row);
    }
  }
}

function resetReminderScheduleForRow_(sheet, rowNumber) {
  const headers = getHeaders_(sheet);
  const values = {
    reminder_count: 0,
    last_reminder_at: '',
    escalated_at: ''
  };

  Object.keys(values).forEach(function (header) {
    if (headers.indexOf(header) >= 0) {
      setCellByHeader_(sheet, rowNumber, header, values[header]);
    }
  });
}

function backfillMissingDeadlines_(sheet, headers) {
  const monthCol = headers.indexOf('month') + 1;
  const deadlineCol = headers.indexOf('deadline') + 1;
  if (monthCol <= 0 || deadlineCol <= 0 || sheet.getLastRow() < 2) {
    return;
  }

  const rowCount = sheet.getLastRow() - 1;
  const months = sheet.getRange(2, monthCol, rowCount, 1).getValues();
  const deadlines = sheet.getRange(2, deadlineCol, rowCount, 1).getValues();
  let changed = false;

  for (let i = 0; i < rowCount; i++) {
    const month = String(months[i][0] || '').trim();
    if (!deadlines[i][0] && isValidMonth_(month)) {
      deadlines[i][0] = getDeadlineForReportingMonth_(month);
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(2, deadlineCol, rowCount, 1).setValues(deadlines);
  }
}

function applyConditionalFormatting_(sheet) {
  const headers = getHeaders_(sheet);
  const allCompleteCol = columnToLetter_(headers.indexOf('all_complete') + 1);
  const daysOverdueCol = columnToLetter_(headers.indexOf('days_overdue') + 1);
  const escalatedCol = columnToLetter_(headers.indexOf('escalated_at') + 1);
  const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + allCompleteCol + '2=TRUE')
      .setBackground('#d9ead3')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + escalatedCol + '2<>"",$' + allCompleteCol + '2=FALSE)')
      .setBackground('#fce5cd')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + daysOverdueCol + '2>0,$' + allCompleteCol + '2=FALSE)')
      .setBackground('#f4cccc')
      .setRanges([range])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}

function findSubmissionRow_(companyName, month) {
  const sheet = getTrackerSheet_();
  const data = getRecords_(sheet);
  const normalizedCompany = normalizeText_(companyName);
  for (let i = 0; i < data.length; i++) {
    if (normalizeText_(data[i].record.company_name) === normalizedCompany && String(data[i].record.month) === String(month)) {
      return data[i].rowNumber;
    }
  }
  return -1;
}

function getRecordAtRow_(sheet, rowNumber) {
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const record = {};
  headers.forEach(function (header, index) {
    record[header] = values[index];
  });
  return record;
}

function getRecords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function (row, index) {
    const record = {};
    headers.forEach(function (header, headerIndex) {
      record[header] = row[headerIndex];
    });
    return {
      rowNumber: index + 2,
      record: record
    };
  });
}

function getLatestCompanyProfiles_() {
  const sheet = getTrackerSheet_();
  const records = getRecords_(sheet);
  const byCompany = {};

  records.forEach(function (entry) {
    const record = entry.record;
    if (!record.company_name) {
      return;
    }
    const key = normalizeText_(record.company_name);
    if (!byCompany[key] || String(record.month) > String(byCompany[key].month)) {
      byCompany[key] = {
        month: record.month,
        company_name: record.company_name,
        ceo_name: record.ceo_name,
        ceo_email: record.ceo_email,
        board_member_name: record.board_member_name,
        board_member_email: record.board_member_email
      };
    }
  });

  return Object.keys(byCompany).map(function (key) {
    return byCompany[key];
  });
}

function getCompanyNames_() {
  return getLatestCompanyProfiles_()
    .map(function (profile) {
      return profile.company_name;
    })
    .filter(function (name, index, names) {
      return name && names.indexOf(name) === index;
    })
    .sort();
}

function getRootFolder_(createIfMissing) {
  const existingId = getScriptProperty_(PROPERTY_KEYS.ROOT_FOLDER_ID);
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (err) {
      clearScriptProperty_(PROPERTY_KEYS.ROOT_FOLDER_ID);
    }
  }

  const folders = DriveApp.getFoldersByName(getConfigString_('ROOT_FOLDER_NAME'));
  if (folders.hasNext()) {
    const folder = folders.next();
    setScriptProperty_(PROPERTY_KEYS.ROOT_FOLDER_ID, folder.getId());
    return folder;
  }

  if (!createIfMissing) {
    return null;
  }

  const folder = DriveApp.createFolder(getConfigString_('ROOT_FOLDER_NAME'));
  setScriptProperty_(PROPERTY_KEYS.ROOT_FOLDER_ID, folder.getId());
  return folder;
}

function getNamedFolder_(parentFolder, name, createIfMissing) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  if (!createIfMissing) {
    return null;
  }
  return parentFolder.createFolder(name);
}

function getCompanyFolder_(companyName, createIfMissing) {
  return getNamedFolder_(getRootFolder_(createIfMissing), companyName, createIfMissing);
}

function getCompanyMonthFolder_(companyName, month, createIfMissing) {
  const companyFolder = getCompanyFolder_(companyName, createIfMissing);
  if (!companyFolder) {
    return null;
  }
  return getNamedFolder_(companyFolder, month, createIfMissing);
}

function applyCompanyFolderSharing_(companyFolder, company) {
  if (!companyFolder || !company) {
    return;
  }

  const companyName = company.companyName || company.company_name || '';
  applyFolderAccess_(companyFolder, company.ceoEmail || company.ceo_email, getConfigValue_('CEO_FOLDER_ACCESS'), 'CEO', companyName);
  applyFolderAccess_(companyFolder, company.boardMemberEmail || company.board_member_email, getConfigValue_('BOARD_MEMBER_FOLDER_ACCESS'), 'board member', companyName);
}

function applyFolderSharingSettings() {
  const companies = getLatestCompanyProfiles_();
  let updatedCount = 0;

  companies.forEach(function (company) {
    const folder = getCompanyFolder_(company.company_name, false);
    if (!folder) {
      return;
    }
    applyCompanyFolderSharing_(folder, company);
    updatedCount++;
  });

  logEvent_('folder_sharing_settings_applied', '', '', 'Applied folder sharing settings to company folders', {
    companyCount: updatedCount,
    ceoAccess: getConfigValue_('CEO_FOLDER_ACCESS'),
    boardMemberAccess: getConfigValue_('BOARD_MEMBER_FOLDER_ACCESS')
  });
  SpreadsheetApp.getUi().alert('Folder sharing settings applied to ' + updatedCount + ' company folders.');
}

function applyFolderAccess_(folder, email, access, roleLabel, companyName) {
  const recipient = String(email || '').trim();
  if (!recipient) {
    return;
  }

  const normalizedAccess = String(access || 'none').toLowerCase();
  try {
    if (normalizedAccess === 'none') {
      removeFolderAccess_(folder, recipient);
      return;
    }

    if (normalizedAccess === 'viewer') {
      try {
        folder.removeEditor(recipient);
      } catch (removeEditorErr) {
        // The user may not be an editor; continue with viewer access.
      }
      folder.addViewer(recipient);
      return;
    }

    if (normalizedAccess === 'editor') {
      folder.addEditor(recipient);
      return;
    }
  } catch (err) {
    logEvent_('folder_sharing_error', companyName || '', '', 'Could not apply folder sharing setting', {
      role: roleLabel,
      email: recipient,
      access: normalizedAccess,
      folderUrl: folder.getUrl(),
      error: err.message
    });
  }
}

function removeFolderAccess_(folder, email) {
  try {
    folder.removeEditor(email);
  } catch (editorErr) {
    // Ignore when the user is not an editor or cannot be removed.
  }
  try {
    folder.removeViewer(email);
  } catch (viewerErr) {
    // Ignore when the user is not a viewer or cannot be removed.
  }
}

function moveFileToFolder_(file, folder) {
  try {
    file.moveTo(folder);
  } catch (err) {
    folder.addFile(file);
  }
}

function getUploadForm_(throwIfMissing) {
  const formId = getScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID);
  if (!formId) {
    if (throwIfMissing) {
      throw new Error('Upload form is not configured. Run setupUploadForm() first.');
    }
    return null;
  }
  try {
    return FormApp.openById(formId);
  } catch (err) {
    if (throwIfMissing) {
      throw err;
    }
    clearScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID);
    return null;
  }
}

function getFormUrl_() {
  const form = getUploadForm_(false);
  return form ? form.getPublishedUrl() : getConfigString_('FORM_URL');
}

function findFormItemByTitle_(form, title) {
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === title) {
      return items[i];
    }
  }
  return null;
}

function findDocByDisplayName_(displayName) {
  const normalized = normalizeText_(displayName);
  const docs = getRequiredDocs_();
  for (let i = 0; i < docs.length; i++) {
    if (normalizeText_(docs[i].display_name) === normalized) {
      return docs[i];
    }
  }
  return null;
}

function getMissingDocs_(record, docs) {
  return docs.filter(function (doc) {
    return String(record['doc_' + doc.doc_key + '_status'] || '').toLowerCase() !== 'received';
  });
}

function isRecordComplete_(record, docs) {
  return getMissingDocs_(record, docs).length === 0;
}

function calculateDaysOverdue_(deadline, now) {
  if (!deadline) {
    return 0;
  }
  const deadlineDate = asDate_(deadline);
  if (isNaN(deadlineDate.getTime())) {
    return 0;
  }
  const currentDate = now ? asDate_(now) : new Date();
  const startCurrent = startOfDay_(currentDate);
  const startDeadline = startOfDay_(deadlineDate);
  const days = Math.floor((startCurrent.getTime() - startDeadline.getTime()) / 86400000);
  return Math.max(0, days);
}

function getDeadlineForReportingMonth_(month) {
  const parts = parseMonth_(month);
  return new Date(parts.year, parts.monthIndex + 1, getConfigInteger_('TARGET_DAY_OF_MONTH'));
}

function getDefaultReportingMonth_(date) {
  const day = Number(Utilities.formatDate(date, getConfigString_('TIMEZONE'), 'd'));
  if (day <= getConfigInteger_('TARGET_DAY_OF_MONTH')) {
    return formatMonth_(addMonths_(date, -1));
  }
  return formatMonth_(date);
}

function getPreviousMonth_(date) {
  return formatMonth_(addMonths_(date, -1));
}

function parseMonth_(month) {
  if (!isValidMonth_(month)) {
    throw new Error('Invalid month "' + month + '". Use YYYY-MM.');
  }
  const parts = String(month).split('-');
  return {
    year: Number(parts[0]),
    month: Number(parts[1]),
    monthIndex: Number(parts[1]) - 1
  };
}

function isValidMonth_(month) {
  return /^\d{4}-\d{2}$/.test(String(month || ''));
}

function formatMonth_(date) {
  return Utilities.formatDate(date, getConfigString_('TIMEZONE'), 'yyyy-MM');
}

function formatDate_(date) {
  if (!date) {
    return '';
  }
  return Utilities.formatDate(asDate_(date), getConfigString_('TIMEZONE'), 'yyyy-MM-dd');
}

function formatDateTime_(date) {
  if (!date) {
    return '';
  }
  return Utilities.formatDate(asDate_(date), getConfigString_('TIMEZONE'), 'yyyy-MM-dd HH:mm');
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addMonths_(date, monthDelta) {
  return new Date(date.getFullYear(), date.getMonth() + monthDelta, 1);
}

function asDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }
  return new Date(value);
}

function columnToLetter_(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function splitExtensions_(acceptedExtensions) {
  return String(acceptedExtensions || '')
    .split(',')
    .map(function (value) {
      return value.trim().toLowerCase().replace(/^\./, '');
    })
    .filter(Boolean);
}

function getFileExtension_(fileName) {
  const match = String(fileName || '').toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : '';
}

function sanitizeDocKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeText_(value) {
  return String(value || '').trim().toLowerCase();
}

function truncate_(text, maxChars) {
  const value = String(text || '');
  if (value.length <= maxChars) {
    return value;
  }
  return value.slice(0, maxChars) + '\n[TRUNCATED after ' + maxChars + ' characters]';
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function safeJsonStringify_(value) {
  try {
    return JSON.stringify(value || {});
  } catch (err) {
    return JSON.stringify({ error: 'Unable to stringify payload', message: err.message });
  }
}

function logEvent_(eventType, companyName, month, message, payload) {
  const sheet = getLogsSheet_();
  sheet.appendRow([
    new Date(),
    eventType,
    companyName || '',
    month || '',
    message || '',
    safeJsonStringify_(payload || {})
  ]);
}

function getScriptProperty_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setScriptProperty_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}

function clearScriptProperty_(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}

function hasTrigger_(handlerFunction) {
  return ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === handlerFunction;
  });
}

function ensureTimeTrigger_(handlerFunction, label, builderFactory) {
  if (hasTrigger_(handlerFunction)) {
    return;
  }
  builderFactory().create();
  logEvent_('trigger_installed', '', '', 'Installed ' + label + ' trigger', {
    handlerFunction: handlerFunction
  });
}

function replaceTimeTrigger_(handlerFunction, label, builderFactory) {
  deleteTriggersForHandler_(handlerFunction);
  builderFactory().create();
  logEvent_('trigger_installed', '', '', 'Installed ' + label + ' trigger', {
    handlerFunction: handlerFunction,
    replacedExisting: true
  });
}

function deleteTriggersForHandler_(handlerFunction) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handlerFunction) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function appendNote_(sheet, rowNumber, note) {
  const existing = String(getCellByHeader_(sheet, rowNumber, 'notes') || '').trim();
  const stamped = '[' + formatDateTime_(new Date()) + '] ' + note;
  setCellByHeader_(sheet, rowNumber, 'notes', existing ? existing + '\n' + stamped : stamped);
}
