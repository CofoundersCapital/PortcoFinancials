const CONFIG = {
  ROOT_FOLDER_NAME: 'Portfolio Reporting',
  TRACKER_FILENAME: 'Submission Tracker',
  TRACKER_SHEET_NAME: 'Submissions',
  LOGS_SHEET_NAME: 'Logs',
  REQUIRED_DOCS_SHEET_NAME: 'Required Documents Checklist',
  FORM_TITLE: 'Monthly Financial Submission',
  INTAKE_EMAIL: 'reporting@cfc.com',
  GMAIL_LOOKBACK_DAYS: 90,
  GMAIL_SEARCH_BATCH_SIZE: 50,
  GMAIL_PROCESSED_LABEL: 'CFC_Reporting_Processed',
  GMAIL_NEEDS_REVIEW_LABEL: 'CFC_Reporting_Needs_Review',
  TARGET_DAY_OF_MONTH: 10,
  REMINDER_INTERVAL_DAYS: 3,
  ESCALATION_THRESHOLD_DAYS: 10,
  REESCALATION_INTERVAL_DAYS: 7,
  ANTHROPIC_MODEL: 'claude-opus-4-7',
  ANTHROPIC_ENDPOINT: 'https://api.anthropic.com/v1/messages',
  ANTHROPIC_VERSION: '2023-06-01',
  FORM_URL: '',
  CFC_TEAM_EMAIL: 'team@cfc.com',
  TIMEZONE: 'America/New_York',
  MAX_EXTRACTED_CHARS_PER_FILE: 60000,
  DEFAULT_DOCS: [
    {
      doc_key: 'financials',
      display_name: 'Monthly Financial Statements',
      accepted_extensions: 'pdf, xlsx'
    },
    {
      doc_key: 'model',
      display_name: 'Updated Financial Model',
      accepted_extensions: 'xlsx'
    },
    {
      doc_key: 'forecast',
      display_name: 'Forecast / Budget Update',
      accepted_extensions: 'xlsx, pdf'
    }
  ]
};

const PROPERTY_KEYS = {
  ROOT_FOLDER_ID: 'CFC_ROOT_FOLDER_ID',
  TRACKER_SPREADSHEET_ID: 'CFC_TRACKER_SPREADSHEET_ID',
  UPLOAD_FORM_ID: 'CFC_UPLOAD_FORM_ID',
  FLASH_TEMPLATE_ID: 'CFC_FLASH_TEMPLATE_ID',
  ANTHROPIC_API_KEY: 'ANTHROPIC_API_KEY'
};

const BASE_SUBMISSION_HEADERS = [
  'month',
  'company_name',
  'ceo_name',
  'ceo_email',
  'board_member_name',
  'board_member_email',
  'deadline'
];

const TRAILING_SUBMISSION_HEADERS = [
  'all_complete',
  'days_overdue',
  'reminder_count',
  'last_reminder_at',
  'escalated_at',
  'flash_report_url',
  'notes'
];

const LOG_HEADERS = [
  'timestamp',
  'event_type',
  'company_name',
  'month',
  'message',
  'payload_json'
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CFC Reporting')
    .addItem('Run setup', 'setupTracker')
    .addItem('Create/update upload form', 'setupUploadForm')
    .addItem('Install triggers', 'installTriggers')
    .addSeparator()
    .addItem('Onboard company from prompts', 'onboardCompanyFromPrompt')
    .addItem('Generate flash report for selected row', 'generateFlashReportForActiveRow')
    .addSeparator()
    .addItem('Run reminder sweep now', 'dailyReminderSweep')
    .addItem('Run email inbox watcher now', 'gmailInboxWatcher')
    .addItem('Run Drive folder watcher now', 'driveFolderWatcher')
    .addToUi();
}

function installTriggers() {
  ensureTimeTrigger_('dailyReminderSweep', 'daily', function () {
    return ScriptApp.newTrigger('dailyReminderSweep')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .inTimezone(CONFIG.TIMEZONE);
  });

  ensureTimeTrigger_('monthlyRolloverJob', 'monthly', function () {
    return ScriptApp.newTrigger('monthlyRolloverJob')
      .timeBased()
      .onMonthDay(1)
      .atHour(6)
      .inTimezone(CONFIG.TIMEZONE);
  });

  ensureTimeTrigger_('driveFolderWatcher', 'watcher', function () {
    return ScriptApp.newTrigger('driveFolderWatcher')
      .timeBased()
      .everyMinutes(30);
  });

  ensureTimeTrigger_('gmailInboxWatcher', 'email watcher', function () {
    return ScriptApp.newTrigger('gmailInboxWatcher')
      .timeBased()
      .everyMinutes(30);
  });

  const formId = getScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID);
  if (formId && !hasTrigger_('onFormSubmit')) {
    ScriptApp.newTrigger('onFormSubmit')
      .forForm(FormApp.openById(formId))
      .onFormSubmit()
      .create();
    logEvent_('trigger_installed', '', '', 'Installed form submit trigger', { formId: formId });
  }

  logEvent_('triggers_checked', '', '', 'Trigger installation complete', {});
}

function getConfigForDebug() {
  return {
    config: CONFIG,
    properties: {
      rootFolderId: getScriptProperty_(PROPERTY_KEYS.ROOT_FOLDER_ID),
      trackerSpreadsheetId: getScriptProperty_(PROPERTY_KEYS.TRACKER_SPREADSHEET_ID),
      uploadFormId: getScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID),
      flashTemplateId: getScriptProperty_(PROPERTY_KEYS.FLASH_TEMPLATE_ID)
    }
  };
}
