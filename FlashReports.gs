function generateFlashReport(companyName, month) {
  if (!isFeatureEnabled_('FLASH_REPORT_GENERATION_ENABLED')) {
    throw new Error('Flash report generation is disabled in Master Config.');
  }

  if (!companyName || !isValidMonth_(month)) {
    throw new Error('Usage: generateFlashReport(companyName, month). Month must use YYYY-MM.');
  }

  const folder = getCompanyMonthFolder_(companyName, month, false);
  if (!folder) {
    throw new Error('No folder found for ' + companyName + ' / ' + month);
  }

  const sourceFiles = listReportSourceFiles_(folder);
  if (sourceFiles.length === 0) {
    throw new Error('No source files found in ' + folder.getUrl());
  }

  const extractedFiles = sourceFiles.map(function (file) {
    return {
      name: file.getName(),
      text: extractFileText_(file)
    };
  });

  const prompt = buildOpenAIExtractionPrompt_(companyName, month, extractedFiles);
  const extraction = openaiExtract(prompt);
  const reportFile = fillDocTemplate_(companyName, month, extraction, folder);
  const rowNumber = findSubmissionRow_(companyName, month);

  if (rowNumber > 0) {
    setCellByHeader_(getTrackerSheet_(), rowNumber, 'flash_report_url', reportFile.getUrl());
  }

  notifyCfcFlashReportReady_(companyName, month, extraction, reportFile.getUrl());

  logEvent_('flash_report_generated', companyName, month, 'Draft flash report generated', {
    reportUrl: reportFile.getUrl(),
    extractionConfidence: extraction.extraction_confidence
  });

  return reportFile.getUrl();
}

function generateFlashReportForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== getConfigString_('TRACKER_SHEET_NAME')) {
    SpreadsheetApp.getUi().alert('Select a row in the ' + getConfigString_('TRACKER_SHEET_NAME') + ' sheet first.');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a company submission row first.');
    return;
  }

  const record = getRecordAtRow_(sheet, row);
  try {
    const url = generateFlashReport(record.company_name, record.month);
    SpreadsheetApp.getUi().alert('Draft flash report created:\n\n' + url);
  } catch (err) {
    SpreadsheetApp.getUi().alert(err.message);
  }
}

function handleSubmissionComplete_(record) {
  sendCompletionNotifications_(record);
  maybeGenerateFlashReportOnComplete_(record);
}

function maybeGenerateFlashReportOnComplete_(record) {
  if (!isFeatureEnabled_('AUTO_GENERATE_FLASH_REPORT_ON_COMPLETE')) {
    return;
  }

  if (!isFeatureEnabled_('FLASH_REPORT_GENERATION_ENABLED')) {
    logEvent_('auto_flash_report_generation_skipped', record.company_name, record.month, 'Auto flash report skipped because generation is disabled', {});
    return;
  }

  if (record.flash_report_url) {
    return;
  }

  try {
    generateFlashReport(record.company_name, record.month);
  } catch (err) {
    logEvent_('auto_flash_report_generation_failed', record.company_name, record.month, 'Auto flash report generation failed', {
      error: err.message
    });
  }
}

function listReportSourceFiles_(folder) {
  const files = [];
  const iterator = folder.getFiles();
  while (iterator.hasNext()) {
    const file = iterator.next();
    if (isSupportedExtractionFile_(file)) {
      files.push(file);
    }
  }
  return files;
}

function isSupportedExtractionFile_(file) {
  const mimeType = file.getMimeType();
  const extension = getFileExtension_(file.getName());
  return mimeType === MimeType.PDF
    || mimeType === MimeType.GOOGLE_DOCS
    || mimeType === MimeType.GOOGLE_SHEETS
    || mimeType === MimeType.CSV
    || mimeType === MimeType.PLAIN_TEXT
    || extension === 'xlsx'
    || extension === 'xls'
    || extension === 'csv'
    || extension === 'pdf'
    || extension === 'txt';
}

function extractFileText_(file) {
  const mimeType = file.getMimeType();
  const extension = getFileExtension_(file.getName());

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    return truncate_(spreadsheetToCsvText_(file.getId()), getConfigInteger_('MAX_EXTRACTED_CHARS_PER_FILE'));
  }

  if (extension === 'xlsx' || extension === 'xls' || mimeType === MimeType.MICROSOFT_EXCEL) {
    return truncate_(excelToCsvText_(file), getConfigInteger_('MAX_EXTRACTED_CHARS_PER_FILE'));
  }

  if (mimeType === MimeType.PDF || extension === 'pdf') {
    return truncate_(pdfToText_(file), getConfigInteger_('MAX_EXTRACTED_CHARS_PER_FILE'));
  }

  if (mimeType === MimeType.GOOGLE_DOCS) {
    return truncate_(DocumentApp.openById(file.getId()).getBody().getText(), getConfigInteger_('MAX_EXTRACTED_CHARS_PER_FILE'));
  }

  if (mimeType === MimeType.CSV || mimeType === MimeType.PLAIN_TEXT || extension === 'csv' || extension === 'txt') {
    return truncate_(file.getBlob().getDataAsString(), getConfigInteger_('MAX_EXTRACTED_CHARS_PER_FILE'));
  }

  return '[Unsupported file type: ' + file.getName() + ' / ' + mimeType + ']';
}

function extractFileTextForClassification_(file) {
  const mimeType = file.getMimeType();
  const extension = getFileExtension_(file.getName());
  const maxChars = getConfigInteger_('OPENAI_CLASSIFICATION_MAX_CHARS') || getConfigInteger_('MAX_EXTRACTED_CHARS_PER_FILE');

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    return spreadsheetToCsvTextWithSheetBudget_(file.getId(), maxChars);
  }

  if (extension === 'xlsx' || extension === 'xls' || mimeType === MimeType.MICROSOFT_EXCEL) {
    return excelToCsvTextWithSheetBudget_(file, maxChars);
  }

  return truncate_(extractFileText_(file), maxChars);
}

function excelToCsvText_(file) {
  const temp = Drive.Files.copy({
    title: 'TEMP converted - ' + file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  }, file.getId(), {
    convert: true
  });

  try {
    return spreadsheetToCsvText_(temp.id);
  } finally {
    DriveApp.getFileById(temp.id).setTrashed(true);
  }
}

function excelToCsvTextWithSheetBudget_(file, maxChars) {
  const temp = Drive.Files.copy({
    title: 'TEMP converted - ' + file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  }, file.getId(), {
    convert: true
  });

  try {
    return spreadsheetToCsvTextWithSheetBudget_(temp.id, maxChars);
  } finally {
    DriveApp.getFileById(temp.id).setTrashed(true);
  }
}

function spreadsheetToCsvText_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  return ss.getSheets().map(function (sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    return [
      '--- SHEET: ' + sheet.getName() + ' ---',
      values.map(function (row) {
        return row.map(csvEscape_).join(',');
      }).join('\n')
    ].join('\n');
  }).join('\n\n');
}

function spreadsheetToCsvTextWithSheetBudget_(spreadsheetId, maxChars) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();
  if (sheets.length === 0) {
    return '';
  }

  const separator = '\n\n';
  const separatorBudget = separator.length * Math.max(0, sheets.length - 1);
  const availableChars = Math.max(1, Number(maxChars || 0) - separatorBudget);
  const charsPerSheet = Math.max(1, Math.floor(availableChars / sheets.length));

  return sheets.map(function (sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    const section = [
      '--- SHEET: ' + sheet.getName() + ' ---',
      values.map(function (row) {
        return row.map(csvEscape_).join(',');
      }).join('\n')
    ].join('\n');
    return truncateToBudget_(section, charsPerSheet);
  }).join(separator);
}

function truncateToBudget_(text, maxChars) {
  const value = String(text || '');
  const limit = Math.max(1, Number(maxChars || 0));
  if (value.length <= limit) {
    return value;
  }

  const marker = '\n[TRUNCATED after ' + limit + ' characters]';
  if (limit <= marker.length) {
    return value.slice(0, limit);
  }
  return value.slice(0, limit - marker.length) + marker;
}

function pdfToText_(file) {
  const temp = Drive.Files.copy({
    title: 'TEMP OCR - ' + file.getName(),
    mimeType: MimeType.GOOGLE_DOCS
  }, file.getId(), {
    ocr: true,
    convert: true
  });

  try {
    return DocumentApp.openById(temp.id).getBody().getText();
  } finally {
    DriveApp.getFileById(temp.id).setTrashed(true);
  }
}

function buildOpenAIExtractionPrompt_(companyName, month, extractedFiles) {
  const sourceText = extractedFiles.map(function (file) {
    return [
      '--- FILE: ' + file.name + ' ---',
      file.text
    ].join('\n');
  }).join('\n\n');

  return [
    'You are extracting structured financial data from a portfolio company monthly submission.',
    '',
    'Company: ' + companyName,
    'Reporting month: ' + month,
    '',
    'Source files (text-extracted, may be messy):',
    sourceText,
    '',
    'Return a JSON object matching the configured schema. Use null for any field you cannot determine confidently. Do not guess.',
    '',
    '{',
    '  "revenue_month": number | null,',
    '  "revenue_currency": "USD" | string,',
    '  "gross_margin_pct": number | null,',
    '  "ebitda_month": number | null,',
    '  "net_income_month": number | null,',
    '  "cash_balance": number | null,',
    '  "burn_rate_monthly": number | null,',
    '  "runway_months": number | null,',
    '  "budget_vs_actual_revenue_pct": number | null,',
    '  "forecast_summary": string | null,',
    '  "key_risks": string[],',
    '  "model_assumption_changes": string[],',
    '  "exceptions_for_partner_attention": string[],',
    '  "extraction_confidence": "high" | "medium" | "low",',
    '  "extraction_notes": string',
    '}'
  ].join('\n');
}

function fillDocTemplate_(companyName, month, extraction, destinationFolder) {
  const templateId = getScriptProperty_(PROPERTY_KEYS.FLASH_TEMPLATE_ID);
  if (!templateId) {
    throw new Error('Flash report template is not configured. Run setupTracker() first.');
  }

  const reportName = 'Flash Report - ' + companyName + ' - ' + month;
  trashExistingFileNamed_(destinationFolder, reportName, '');
  const reportFile = DriveApp.getFileById(templateId).makeCopy(reportName, destinationFolder);
  const doc = DocumentApp.openById(reportFile.getId());
  const body = doc.getBody();
  const values = buildFlashReportTemplateValues_(companyName, month, extraction);

  Object.keys(values).forEach(function (key) {
    body.replaceText('{{' + escapeRegExp_(key) + '}}', values[key]);
  });

  doc.saveAndClose();
  return reportFile;
}

function buildFlashReportTemplateValues_(companyName, month, extraction) {
  return {
    company_name: companyName,
    month: month,
    revenue_month: formatMoney_(extraction.revenue_month, extraction.revenue_currency),
    gross_margin_pct: formatPercent_(extraction.gross_margin_pct),
    ebitda_month: formatMoney_(extraction.ebitda_month, extraction.revenue_currency),
    net_income_month: formatMoney_(extraction.net_income_month, extraction.revenue_currency),
    cash_balance: formatMoney_(extraction.cash_balance, extraction.revenue_currency),
    burn_rate_monthly: formatMoney_(extraction.burn_rate_monthly, extraction.revenue_currency),
    runway_months: formatNullable_(extraction.runway_months),
    budget_vs_actual_revenue_pct: formatPercent_(extraction.budget_vs_actual_revenue_pct),
    forecast_summary: formatNullable_(extraction.forecast_summary),
    key_risks: formatBullets_(extraction.key_risks),
    model_assumption_changes: formatBullets_(extraction.model_assumption_changes),
    exceptions_for_partner_attention: formatBullets_(extraction.exceptions_for_partner_attention),
    extraction_confidence: formatNullable_(extraction.extraction_confidence),
    extraction_notes: formatNullable_(extraction.extraction_notes)
  };
}

function notifyCfcFlashReportReady_(companyName, month, extraction, reportUrl) {
  if (!isFeatureEnabled_('NOTIFY_TEAM_ON_FLASH_REPORT_READY')) {
    return;
  }

  const data = {
    company_name: companyName,
    month: month,
    extraction_confidence: extraction.extraction_confidence || '',
    flash_report_url: reportUrl,
    tracker_url: getTrackerSpreadsheet_().getUrl()
  };
  sendTemplatedEmail_(getConfigString_('CFC_TEAM_EMAIL'), EMAIL_TEMPLATES.CFC_REVIEW_FLASH_REPORT, data);
}

function csvEscape_(value) {
  const text = String(value || '');
  if (/[",\n]/.test(text)) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

function formatMoney_(value, currency) {
  if (value === null || value === undefined || value === '') {
    return 'Not extracted';
  }
  const number = Number(value);
  if (isNaN(number)) {
    return String(value);
  }
  const prefix = String(currency || 'USD') + ' ';
  return prefix + number.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

function formatPercent_(value) {
  if (value === null || value === undefined || value === '') {
    return 'Not extracted';
  }
  const number = Number(value);
  if (isNaN(number)) {
    return String(value);
  }
  return number.toLocaleString('en-US', { maximumFractionDigits: 1 }) + '%';
}

function formatNullable_(value) {
  if (value === null || value === undefined || value === '') {
    return 'Not extracted';
  }
  return String(value);
}

function formatBullets_(values) {
  if (!values || values.length === 0) {
    return 'Not extracted';
  }
  if (!Array.isArray(values)) {
    return String(values);
  }
  return values.map(function (value) {
    return '- ' + value;
  }).join('\n');
}
