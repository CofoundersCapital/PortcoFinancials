function setupTracker() {
  const ss = getOrCreateTrackerSpreadsheet_();
  setScriptProperty_(PROPERTY_KEYS.TRACKER_SPREADSHEET_ID, ss.getId());

  const rootFolder = getRootFolder_(true);
  const configFolder = getNamedFolder_(rootFolder, '_config', true);
  const trackerFolder = getNamedFolder_(rootFolder, '_tracker', true);

  moveFileToFolder_(DriveApp.getFileById(ss.getId()), trackerFolder);

  setupMasterConfigSheet_(ss);
  setupRequiredDocsSheet_(ss);
  setupSubmissionsSheet_(ss);
  setupLogsSheet_(ss);
  setupClassificationLogSheet_(ss);
  ensureFlashReportTemplate_(configFolder);
  setupEmailIntake_();

  logEvent_('setup_complete', '', '', 'Tracker setup completed', {
    spreadsheetUrl: ss.getUrl(),
    rootFolderUrl: rootFolder.getUrl(),
    intakeEmail: getConfigString_('INTAKE_EMAIL')
  });

  return {
    spreadsheetUrl: ss.getUrl(),
    rootFolderUrl: rootFolder.getUrl(),
    intakeEmail: getConfigString_('INTAKE_EMAIL')
  };
}

function setupUploadForm() {
  const docs = getRequiredDocs_();
  let form = getUploadForm_(false);

  if (!form) {
    form = FormApp.create(getConfigString_('FORM_TITLE'));
    setScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID, form.getId());
    moveFileToFolder_(DriveApp.getFileById(form.getId()), getRootFolder_(true));
  }

  form.setTitle(getConfigString_('FORM_TITLE'));
  form.setDescription('Submit monthly financial materials for Cofounders Capital reporting.');
  form.setAcceptingResponses(true);
  form.setCollectEmail(true);

  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i--) {
    form.deleteItem(items[i]);
  }

  const companyItem = form.addListItem()
    .setTitle('Company')
    .setRequired(true);
  companyItem.setChoices(buildCompanyChoices_(form, companyItem));

  form.addTextItem()
    .setTitle('Reporting Month (YYYY-MM)')
    .setHelpText('Example: 2026-05')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation()
      .requireTextMatchesPattern('^\\d{4}-\\d{2}$')
      .setHelpText('Use YYYY-MM format, for example 2026-05.')
      .build());

  docs.forEach(function (doc) {
    form.addParagraphTextItem()
      .setTitle(doc.display_name)
      .setHelpText('Paste the Google Drive file link after uploading this file to your shared company folder. Accepted extensions: ' + doc.accepted_extensions)
      .setRequired(false);
  });

  form.addParagraphTextItem()
    .setTitle('Notes')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Other / Supporting Documents')
    .setRequired(false)
    .setHelpText('Optional: paste Google Drive links for supporting materials requested by CFC.');

  setScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID, form.getId());
  logEvent_('form_setup', '', '', 'Upload form created or refreshed', {
    formUrl: form.getPublishedUrl()
  });

  return form.getPublishedUrl();
}

function updateFormCompanyChoices() {
  const form = getUploadForm_(false);
  if (!form) {
    return;
  }

  const companyItem = findFormItemByTitle_(form, 'Company');
  if (!companyItem) {
    return;
  }

  const listItem = companyItem.asListItem();
  listItem.setChoices(buildCompanyChoices_(form, listItem));
}

function onboardCompany(companyInput) {
  const company = normalizeCompanyInput_(companyInput);
  const ss = getTrackerSpreadsheet_();
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  const month = company.reportingMonth || getDefaultReportingMonth_(new Date());
  const deadline = company.deadline || getDeadlineForReportingMonth_(month);

  ensureTrackerColumns_(sheet, docs);

  const companyFolder = getCompanyFolder_(company.companyName, true);
  applyCompanyFolderSharing_(companyFolder, company);
  getCompanyMonthFolder_(company.companyName, month, true);

  const existingRow = findSubmissionRow_(company.companyName, month);
  if (existingRow > 0) {
    updateCompanyContactFields_(existingRow, company);
    applySubmissionFormulas_(sheet, existingRow, docs);
    logEvent_('company_onboarded_existing', company.companyName, month, 'Updated existing company row', company);
  } else {
    appendSubmissionRow_(sheet, docs, {
      month: month,
      company_name: company.companyName,
      ceo_name: company.ceoName,
      ceo_email: company.ceoEmail,
      board_member_name: company.boardMemberName,
      board_member_email: company.boardMemberEmail,
      deadline: deadline
    });
    logEvent_('company_onboarded', company.companyName, month, 'Created company folder and tracker row', company);
  }

  updateFormCompanyChoices();

  return {
    trackerUrl: ss.getUrl(),
    companyFolderUrl: companyFolder.getUrl(),
    month: month,
    deadline: formatDate_(deadline)
  };
}

function onboardCompanyFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const fields = [
    ['companyName', 'Company name'],
    ['ceoName', 'CEO name'],
    ['ceoEmail', 'CEO email'],
    ['boardMemberName', 'CFC board member name'],
    ['boardMemberEmail', 'CFC board member email']
  ];
  const input = {};

  for (let i = 0; i < fields.length; i++) {
    const response = ui.prompt('Onboard company', fields[i][1], ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    input[fields[i][0]] = response.getResponseText().trim();
  }

  const result = onboardCompany(input);
  ui.alert('Company onboarded for ' + result.month + '.\n\nFolder: ' + result.companyFolderUrl);
}

function monthlyRolloverJob() {
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  const month = getPreviousMonth_(new Date());
  const companies = getLatestCompanyProfiles_();
  let createdCount = 0;

  companies.forEach(function (company) {
    if (findSubmissionRow_(company.company_name, month) > 0) {
      return;
    }

    appendSubmissionRow_(sheet, docs, {
      month: month,
      company_name: company.company_name,
      ceo_name: company.ceo_name,
      ceo_email: company.ceo_email,
      board_member_name: company.board_member_name,
      board_member_email: company.board_member_email,
      deadline: getDeadlineForReportingMonth_(month)
    });
    const companyFolder = getCompanyFolder_(company.company_name, true);
    applyCompanyFolderSharing_(companyFolder, company);
    getCompanyMonthFolder_(company.company_name, month, true);
    createdCount++;
  });

  updateFormCompanyChoices();
  logEvent_('monthly_rollover', '', month, 'Monthly rollover created rows', { createdCount: createdCount });
}

function setupMasterConfigSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, CONFIG.MASTER_CONFIG_SHEET_NAME);
  const existingValuesByKey = getExistingMasterConfigValues_(sheet);
  const rows = MASTER_CONFIG_DEFINITIONS.map(function (definition) {
    const existing = existingValuesByKey[definition.key];
    const value = existing
      ? existing.value
      : serializeConfigValue_(definition.defaultValue, definition.valueType);
    return [
      definition.category,
      definition.key,
      definition.displayName,
      value,
      serializeConfigValue_(definition.defaultValue, definition.valueType),
      definition.valueType,
      definition.allowedValues,
      definition.description
    ];
  });

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.clear();
  sheet.getRange(1, 1, 1, MASTER_CONFIG_HEADERS.length).setValues([MASTER_CONFIG_HEADERS]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, MASTER_CONFIG_HEADERS.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, MASTER_CONFIG_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(1, 1, sheet.getMaxRows(), MASTER_CONFIG_HEADERS.length).createFilter();
  sheet.getRange(2, 4, Math.max(1, rows.length), 1).setBackground('#fff2cc');
  applyMasterConfigValidation_(sheet, rows.length);
  sheet.autoResizeColumns(1, MASTER_CONFIG_HEADERS.length);
  clearMasterConfigCache_();
}

function refreshMasterConfigSheet() {
  setupMasterConfigSheet_(getTrackerSpreadsheet_());
  SpreadsheetApp.getUi().alert('Master Config refreshed. Existing value cells were preserved.');
}

function getExistingMasterConfigValues_(sheet) {
  const valuesByKey = {};
  if (!sheet || sheet.getLastRow() < 2) {
    return valuesByKey;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (value) {
    return String(value || '').trim();
  });
  const keyIndex = headers.indexOf('setting_key');
  const valueIndex = headers.indexOf('value');
  if (keyIndex < 0 || valueIndex < 0) {
    return valuesByKey;
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  rows.forEach(function (row) {
    const key = String(row[keyIndex] || '').trim();
    if (key) {
      valuesByKey[key] = {
        value: row[valueIndex]
      };
    }
  });
  return valuesByKey;
}

function applyMasterConfigValidation_(sheet, definitionCount) {
  if (definitionCount < 1) {
    return;
  }

  const valueColumn = 4;
  const typeColumn = 6;
  const allowedColumn = 7;
  sheet.getRange(2, typeColumn, definitionCount, 1).setNumberFormat('@');
  sheet.getRange(2, allowedColumn, definitionCount, 1).setNumberFormat('@');

  MASTER_CONFIG_DEFINITIONS.forEach(function (definition, index) {
    const row = index + 2;
    const valueCell = sheet.getRange(row, valueColumn);
    valueCell.clearDataValidations();

    if (definition.valueType === 'boolean') {
      valueCell.insertCheckboxes();
      return;
    }

    if (definition.valueType === 'enum') {
      const choices = splitAllowedValues_(definition.allowedValues);
      if (choices.length > 0) {
        valueCell.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInList(choices, true)
          .setAllowInvalid(false)
          .build());
      }
      return;
    }

    if (definition.valueType === 'integer') {
      valueCell.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setAllowInvalid(false)
        .build());
    }
  });
}

function setupRequiredDocsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, getConfigString_('REQUIRED_DOCS_SHEET_NAME'));
  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([['doc_key', 'display_name', 'accepted_extensions']]);
  const values = CONFIG.DEFAULT_DOCS.map(function (doc) {
    return [doc.doc_key, doc.display_name, doc.accepted_extensions];
  });
  sheet.getRange(2, 1, values.length, 3).setValues(values);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
}

function setupSubmissionsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, getConfigString_('TRACKER_SHEET_NAME'));
  const docs = getRequiredDocs_();
  const headers = buildSubmissionHeaders_(docs);
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).createFilter();
  applyStatusValidation_(sheet, docs);
  applyDeadlineValidation_(sheet, headers);
  applyTrackerNumberFormats_(sheet, headers);
  applyConditionalFormatting_(sheet, docs);
  sheet.autoResizeColumns(1, headers.length);
}

function refreshSubmissionTrackerLayout() {
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  ensureTrackerColumns_(sheet, docs);
  const headers = getHeaders_(sheet);

  backfillMissingDeadlines_(sheet, headers);
  applyStatusValidation_(sheet, docs);
  applyDeadlineValidation_(sheet, headers);
  applyTrackerNumberFormats_(sheet, headers);
  applyConditionalFormatting_(sheet);

  for (let row = 2; row <= sheet.getLastRow(); row++) {
    applySubmissionFormulas_(sheet, row, docs);
  }

  sheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert('Tracker columns, deadline validation, and formulas refreshed.');
}

function setupLogsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, getConfigString_('LOGS_SHEET_NAME'));
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  }
  sheet.autoResizeColumns(1, LOG_HEADERS.length);
}

function setupClassificationLogSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, getConfigString_('CLASSIFICATION_LOG_SHEET_NAME'));
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank()) {
    sheet.getRange(1, 1, 1, CLASSIFICATION_LOG_HEADERS.length).setValues([CLASSIFICATION_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CLASSIFICATION_LOG_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.getRange(1, 1, sheet.getMaxRows(), CLASSIFICATION_LOG_HEADERS.length).createFilter();
  }
  sheet.autoResizeColumns(1, CLASSIFICATION_LOG_HEADERS.length);
}

function ensureFlashReportTemplate_(configFolder) {
  const existingTemplateId = getScriptProperty_(PROPERTY_KEYS.FLASH_TEMPLATE_ID);
  if (existingTemplateId) {
    try {
      DriveApp.getFileById(existingTemplateId);
      return existingTemplateId;
    } catch (err) {
      clearScriptProperty_(PROPERTY_KEYS.FLASH_TEMPLATE_ID);
    }
  }

  const doc = DocumentApp.create('Flash Report Template');
  const body = doc.getBody();
  body.clear();
  body.appendParagraph('DRAFT - CFC Monthly Flash Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('DRAFT - review extraction_confidence and extraction_notes before use.').setBold(true);
  body.appendParagraph('Company: {{company_name}}');
  body.appendParagraph('Reporting month: {{month}}');
  body.appendParagraph('');
  body.appendParagraph('Financial Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Revenue: {{revenue_month}}');
  body.appendParagraph('Gross margin: {{gross_margin_pct}}');
  body.appendParagraph('EBITDA: {{ebitda_month}}');
  body.appendParagraph('Net income: {{net_income_month}}');
  body.appendParagraph('Cash balance: {{cash_balance}}');
  body.appendParagraph('Burn rate: {{burn_rate_monthly}}');
  body.appendParagraph('Runway: {{runway_months}}');
  body.appendParagraph('Budget versus actual revenue: {{budget_vs_actual_revenue_pct}}');
  body.appendParagraph('');
  body.appendParagraph('Forecast Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{forecast_summary}}');
  body.appendParagraph('');
  body.appendParagraph('Key Risks / Changes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{key_risks}}');
  body.appendParagraph('');
  body.appendParagraph('Model Assumption Changes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{model_assumption_changes}}');
  body.appendParagraph('');
  body.appendParagraph('Partner Attention').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{exceptions_for_partner_attention}}');
  body.appendParagraph('');
  body.appendParagraph('Extraction Review').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Confidence: {{extraction_confidence}}');
  body.appendParagraph('Notes: {{extraction_notes}}');
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  moveFileToFolder_(file, configFolder);
  setScriptProperty_(PROPERTY_KEYS.FLASH_TEMPLATE_ID, doc.getId());
  return doc.getId();
}

function appendSubmissionRow_(sheet, docs, rowData) {
  const headers = buildSubmissionHeaders_(docs);
  ensureTrackerColumns_(sheet, docs);
  const rowNumber = Math.max(sheet.getLastRow() + 1, 2);
  const values = headers.map(function (header) {
    if (rowData.hasOwnProperty(header)) {
      return rowData[header];
    }
    if (header === 'deadline' && isValidMonth_(rowData.month)) {
      return getDeadlineForReportingMonth_(rowData.month);
    }
    if (header.indexOf('doc_') === 0 && header.slice(-7) === '_status') {
      return 'missing';
    }
    if (header === 'reminder_count') {
      return 0;
    }
    if (header === 'all_complete') {
      return buildAllCompleteFormula_(rowNumber, docs, headers);
    }
    if (header === 'days_overdue') {
      return buildDaysOverdueFormula_(rowNumber, headers);
    }
    return '';
  });
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  applySubmissionFormulas_(sheet, rowNumber, docs);
}

function updateCompanyContactFields_(rowNumber, company) {
  const sheet = getTrackerSheet_();
  const updates = {
    ceo_name: company.ceoName,
    ceo_email: company.ceoEmail,
    board_member_name: company.boardMemberName,
    board_member_email: company.boardMemberEmail
  };

  Object.keys(updates).forEach(function (header) {
    setCellByHeader_(sheet, rowNumber, header, updates[header]);
  });
}

function normalizeCompanyInput_(input) {
  if (!input || typeof input !== 'object') {
    throw new Error('Pass an object: onboardCompany({ companyName, ceoName, ceoEmail, boardMemberName, boardMemberEmail })');
  }

  const company = {
    companyName: String(input.companyName || input.company_name || '').trim(),
    ceoName: String(input.ceoName || input.ceo_name || '').trim(),
    ceoEmail: String(input.ceoEmail || input.ceo_email || '').trim(),
    boardMemberName: String(input.boardMemberName || input.board_member_name || '').trim(),
    boardMemberEmail: String(input.boardMemberEmail || input.board_member_email || '').trim(),
    reportingMonth: input.reportingMonth || input.month || '',
    deadline: input.deadline || null
  };

  ['companyName', 'ceoName', 'ceoEmail', 'boardMemberName', 'boardMemberEmail'].forEach(function (field) {
    if (!company[field]) {
      throw new Error('Missing required company field: ' + field);
    }
  });

  if (company.reportingMonth && !isValidMonth_(company.reportingMonth)) {
    throw new Error('reportingMonth must use YYYY-MM format.');
  }

  return company;
}

function applyAllowedFileTypes_(item, acceptedExtensions) {
  const extensions = splitExtensions_(acceptedExtensions);
  const allowedTypes = [];

  if (extensions.indexOf('pdf') >= 0) {
    allowedTypes.push(FormApp.FileType.PDF);
  }
  if (extensions.indexOf('xlsx') >= 0 || extensions.indexOf('xls') >= 0 || extensions.indexOf('csv') >= 0) {
    allowedTypes.push(FormApp.FileType.SPREADSHEET);
  }

  if (allowedTypes.length > 0) {
    if (typeof item.setAllowOnlySpecificFileTypes === 'function') {
      item.setAllowOnlySpecificFileTypes(true);
    }
    item.setAllowedFileTypes(allowedTypes);
  }
}

function buildCompanyChoices_(form, item) {
  const companies = getCompanyNames_();
  if (companies.length === 0) {
    return [item.createChoice('Add companies in the tracker first')];
  }
  return companies.map(function (companyName) {
    return item.createChoice(companyName);
  });
}
