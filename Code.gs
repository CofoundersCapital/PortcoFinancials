const CONFIG = {
  ROOT_FOLDER_NAME: 'Portfolio Reporting',
  TRACKER_FILENAME: 'Submission Tracker',
  MASTER_CONFIG_SHEET_NAME: 'Master Config',
  TRACKER_SHEET_NAME: 'Submissions',
  LOGS_SHEET_NAME: 'Logs',
  CLASSIFICATION_LOG_SHEET_NAME: 'Document Classification Log',
  REQUIRED_DOCS_SHEET_NAME: 'Required Documents Checklist',
  FORM_TITLE: 'Monthly Financial Submission',
  INTAKE_EMAIL: 'cofoundersreporting@gmail.com',
  GMAIL_LOOKBACK_DAYS: 90,
  GMAIL_SEARCH_BATCH_SIZE: 50,
  GMAIL_WATCHER_INTERVAL_MINUTES: 30,
  DRIVE_WATCHER_INTERVAL_MINUTES: 30,
  GMAIL_PROCESSED_LABEL: 'CFC_Reporting_Processed',
  GMAIL_NEEDS_REVIEW_LABEL: 'CFC_Reporting_Needs_Review',
  TARGET_DAY_OF_MONTH: 10,
  REMINDER_INTERVAL_DAYS: 3,
  ESCALATION_THRESHOLD_DAYS: 10,
  REESCALATION_INTERVAL_DAYS: 7,
  OPENAI_MODEL: 'gpt-5.4-nano',
  OPENAI_ENDPOINT: 'https://api.openai.com/v1/responses',
  OPENAI_MAX_OUTPUT_TOKENS: 4096,
  OPENAI_CLASSIFICATION_MAX_CHARS: 12000,
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
  OPENAI_API_KEY: 'OPENAI_API_KEY'
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

const CLASSIFICATION_LOG_HEADERS = [
  'timestamp',
  'source',
  'company_name',
  'month',
  'file_name',
  'file_url',
  'mime_type',
  'extension',
  'doc_key',
  'display_name',
  'accepted_extensions',
  'confidence',
  'llm_match',
  'accepted_for_tracker',
  'reason',
  'unmatched_reason'
];

const MASTER_CONFIG_HEADERS = [
  'category',
  'setting_key',
  'display_name',
  'value',
  'default_value',
  'value_type',
  'allowed_values',
  'description'
];

const MASTER_CONFIG_DEFINITIONS = [
  {
    category: 'Drive and sheets',
    key: 'ROOT_FOLDER_NAME',
    displayName: 'Root folder name',
    defaultValue: CONFIG.ROOT_FOLDER_NAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Top-level Google Drive folder for portfolio reporting files.'
  },
  {
    category: 'Drive and sheets',
    key: 'TRACKER_FILENAME',
    displayName: 'Tracker spreadsheet filename',
    defaultValue: CONFIG.TRACKER_FILENAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Name used when creating a new tracker spreadsheet.'
  },
  {
    category: 'Drive and sheets',
    key: 'MASTER_CONFIG_SHEET_NAME',
    displayName: 'Master config sheet name',
    defaultValue: CONFIG.MASTER_CONFIG_SHEET_NAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Name of this configuration tab. This value is visible for reference; rename with care.'
  },
  {
    category: 'Drive and sheets',
    key: 'TRACKER_SHEET_NAME',
    displayName: 'Submissions sheet name',
    defaultValue: CONFIG.TRACKER_SHEET_NAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Sheet tab containing monthly submission tracker rows.'
  },
  {
    category: 'Drive and sheets',
    key: 'LOGS_SHEET_NAME',
    displayName: 'Logs sheet name',
    defaultValue: CONFIG.LOGS_SHEET_NAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Sheet tab where system events are logged.'
  },
  {
    category: 'Drive and sheets',
    key: 'CLASSIFICATION_LOG_SHEET_NAME',
    displayName: 'Classification log sheet name',
    defaultValue: CONFIG.CLASSIFICATION_LOG_SHEET_NAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Sheet tab where LLM document classification audit rows are written.'
  },
  {
    category: 'Drive and sheets',
    key: 'REQUIRED_DOCS_SHEET_NAME',
    displayName: 'Required docs sheet name',
    defaultValue: CONFIG.REQUIRED_DOCS_SHEET_NAME,
    valueType: 'string',
    allowedValues: '',
    description: 'Sheet tab containing the required document checklist.'
  },
  {
    category: 'Drive and sheets',
    key: 'FORM_URL',
    displayName: 'Fallback upload form URL',
    defaultValue: CONFIG.FORM_URL,
    valueType: 'string',
    allowedValues: '',
    description: 'Optional fallback URL shown in emails when no generated upload form is configured.'
  },
  {
    category: 'Folder sharing',
    key: 'CEO_FOLDER_ACCESS',
    displayName: 'CEO folder access',
    defaultValue: 'editor',
    valueType: 'enum',
    allowedValues: 'none, viewer, editor',
    description: 'Access granted to each CEO on their company reporting folder.'
  },
  {
    category: 'Folder sharing',
    key: 'BOARD_MEMBER_FOLDER_ACCESS',
    displayName: 'CFC board member folder access',
    defaultValue: 'editor',
    valueType: 'enum',
    allowedValues: 'none, viewer, editor',
    description: 'Access granted to the assigned CFC board member on each company reporting folder.'
  },
  {
    category: 'Email intake',
    key: 'INTAKE_EMAIL',
    displayName: 'Intake email',
    defaultValue: CONFIG.INTAKE_EMAIL,
    valueType: 'string',
    allowedValues: '',
    description: 'Mailbox address used in Gmail intake searches and reminder copy.'
  },
  {
    category: 'Email intake',
    key: 'CFC_TEAM_EMAIL',
    displayName: 'CFC team email',
    defaultValue: CONFIG.CFC_TEAM_EMAIL,
    valueType: 'string',
    allowedValues: '',
    description: 'Global CFC team mailbox for operational warnings and review notifications.'
  },
  {
    category: 'Email intake',
    key: 'GMAIL_LOOKBACK_DAYS',
    displayName: 'Gmail lookback days',
    defaultValue: CONFIG.GMAIL_LOOKBACK_DAYS,
    valueType: 'integer',
    allowedValues: '',
    description: 'How many recent days Gmail intake searches for unprocessed attachment emails.'
  },
  {
    category: 'Email intake',
    key: 'GMAIL_SEARCH_BATCH_SIZE',
    displayName: 'Gmail search batch size',
    defaultValue: CONFIG.GMAIL_SEARCH_BATCH_SIZE,
    valueType: 'integer',
    allowedValues: '',
    description: 'Maximum number of Gmail threads processed per watcher run.'
  },
  {
    category: 'Automatic watchers',
    key: 'GMAIL_WATCHER_INTERVAL_MINUTES',
    displayName: 'Gmail watcher interval minutes',
    defaultValue: CONFIG.GMAIL_WATCHER_INTERVAL_MINUTES,
    valueType: 'enum',
    allowedValues: '1, 5, 10, 15, 30',
    description: 'How often the automatic Gmail intake watcher runs. Run Install triggers after changing this.'
  },
  {
    category: 'Automatic watchers',
    key: 'DRIVE_WATCHER_INTERVAL_MINUTES',
    displayName: 'Drive watcher interval minutes',
    defaultValue: CONFIG.DRIVE_WATCHER_INTERVAL_MINUTES,
    valueType: 'enum',
    allowedValues: '1, 5, 10, 15, 30',
    description: 'How often the automatic Drive folder watcher runs. Run Install triggers after changing this.'
  },
  {
    category: 'Email intake',
    key: 'GMAIL_PROCESSED_LABEL',
    displayName: 'Gmail processed label',
    defaultValue: CONFIG.GMAIL_PROCESSED_LABEL,
    valueType: 'string',
    allowedValues: '',
    description: 'Gmail label applied after an intake email is fully processed.'
  },
  {
    category: 'Email intake',
    key: 'GMAIL_NEEDS_REVIEW_LABEL',
    displayName: 'Gmail needs review label',
    defaultValue: CONFIG.GMAIL_NEEDS_REVIEW_LABEL,
    valueType: 'string',
    allowedValues: '',
    description: 'Gmail label applied when an intake email needs manual review.'
  },
  {
    category: 'Forms',
    key: 'FORM_TITLE',
    displayName: 'Upload form title',
    defaultValue: CONFIG.FORM_TITLE,
    valueType: 'string',
    allowedValues: '',
    description: 'Title used when creating or refreshing the upload form.'
  },
  {
    category: 'Deadlines and reminders',
    key: 'TARGET_DAY_OF_MONTH',
    displayName: 'Default deadline day of month',
    defaultValue: CONFIG.TARGET_DAY_OF_MONTH,
    valueType: 'integer',
    allowedValues: '',
    description: 'Default day of the month used when creating submission deadlines.'
  },
  {
    category: 'Deadlines and reminders',
    key: 'FIRST_REMINDER_DAYS_OVERDUE',
    displayName: 'First reminder days overdue',
    defaultValue: 1,
    valueType: 'integer',
    allowedValues: '',
    description: 'Days overdue required before the first CEO reminder can send.'
  },
  {
    category: 'Deadlines and reminders',
    key: 'REMINDER_INTERVAL_DAYS',
    displayName: 'Reminder interval days',
    defaultValue: CONFIG.REMINDER_INTERVAL_DAYS,
    valueType: 'integer',
    allowedValues: '',
    description: 'Minimum days between follow-up CEO reminders.'
  },
  {
    category: 'Deadlines and reminders',
    key: 'ESCALATION_THRESHOLD_DAYS',
    displayName: 'Escalation threshold days',
    defaultValue: CONFIG.ESCALATION_THRESHOLD_DAYS,
    valueType: 'integer',
    allowedValues: '',
    description: 'Days overdue before escalation to the assigned CFC board member.'
  },
  {
    category: 'Deadlines and reminders',
    key: 'REESCALATION_INTERVAL_DAYS',
    displayName: 'Re-escalation interval days',
    defaultValue: CONFIG.REESCALATION_INTERVAL_DAYS,
    valueType: 'integer',
    allowedValues: '',
    description: 'Minimum days between repeated escalation emails.'
  },
  {
    category: 'OpenAI',
    key: 'OPENAI_MODEL',
    displayName: 'OpenAI model',
    defaultValue: CONFIG.OPENAI_MODEL,
    valueType: 'string',
    allowedValues: '',
    description: 'OpenAI model used for classification and flash report extraction.'
  },
  {
    category: 'OpenAI',
    key: 'OPENAI_ENDPOINT',
    displayName: 'OpenAI endpoint',
    defaultValue: CONFIG.OPENAI_ENDPOINT,
    valueType: 'string',
    allowedValues: '',
    description: 'Responses API endpoint. The API key remains in Apps Script properties.'
  },
  {
    category: 'OpenAI',
    key: 'OPENAI_MAX_OUTPUT_TOKENS',
    displayName: 'OpenAI max output tokens',
    defaultValue: CONFIG.OPENAI_MAX_OUTPUT_TOKENS,
    valueType: 'integer',
    allowedValues: '',
    description: 'Maximum output tokens for flash report extraction responses.'
  },
  {
    category: 'OpenAI',
    key: 'OPENAI_CLASSIFICATION_MAX_CHARS',
    displayName: 'Classification max characters',
    defaultValue: CONFIG.OPENAI_CLASSIFICATION_MAX_CHARS,
    valueType: 'integer',
    allowedValues: '',
    description: 'Maximum extracted characters sent to the LLM classifier per file.'
  },
  {
    category: 'OpenAI',
    key: 'MAX_EXTRACTED_CHARS_PER_FILE',
    displayName: 'Flash extraction max characters',
    defaultValue: CONFIG.MAX_EXTRACTED_CHARS_PER_FILE,
    valueType: 'integer',
    allowedValues: '',
    description: 'Maximum extracted characters per source file for flash report generation.'
  },
  {
    category: 'Workflow toggles',
    key: 'EMAIL_INTAKE_ENABLED',
    displayName: 'Email intake enabled',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When FALSE, Gmail intake watcher runs are skipped.'
  },
  {
    category: 'Workflow toggles',
    key: 'DRIVE_WATCHER_ENABLED',
    displayName: 'Drive watcher enabled',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When FALSE, Drive folder watcher runs are skipped.'
  },
  {
    category: 'Workflow toggles',
    key: 'FORM_INTAKE_ENABLED',
    displayName: 'Form intake enabled',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When FALSE, form submissions are logged and ignored.'
  },
  {
    category: 'Workflow toggles',
    key: 'OVERDUE_REMINDERS_ENABLED',
    displayName: 'Overdue reminders enabled',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When FALSE, CEO overdue reminder emails are not sent.'
  },
  {
    category: 'Workflow toggles',
    key: 'ESCALATIONS_ENABLED',
    displayName: 'Escalations enabled',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When FALSE, overdue escalation emails are not sent.'
  },
  {
    category: 'Workflow toggles',
    key: 'FLASH_REPORT_GENERATION_ENABLED',
    displayName: 'Flash report generation enabled',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When FALSE, manual and automatic flash report generation are blocked.'
  },
  {
    category: 'Workflow toggles',
    key: 'AUTO_GENERATE_FLASH_REPORT_ON_COMPLETE',
    displayName: 'Auto-generate flash report on completion',
    defaultValue: false,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'When TRUE, creates a draft flash report after a submission first becomes complete.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_BOARD_ON_DOCUMENT_RECEIVED',
    displayName: 'Notify board member when document received',
    defaultValue: false,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the assigned CFC board member when new required documents are received.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_BOARD_ON_SUBMISSION_COMPLETE',
    displayName: 'Notify board member when all docs submitted',
    defaultValue: false,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the assigned CFC board member when a monthly submission becomes complete.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_FOUNDER_ON_DOCUMENT_RECEIVED',
    displayName: 'Notify founder when document received',
    defaultValue: false,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the CEO after newly received required documents are processed.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_FOUNDER_ON_SUBMISSION_COMPLETE',
    displayName: 'Notify founder when all docs submitted',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the CEO when all required monthly documents are received.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_BOARD_ON_ESCALATION',
    displayName: 'Notify board member on escalation',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the assigned CFC board member when a submission is seriously overdue.'
  },
  {
    category: 'Notifications',
    key: 'CC_FOUNDER_ON_ESCALATION',
    displayName: 'CC founder on escalation',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'CC the CEO on board member escalation emails.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_TEAM_ON_FORM_WARNING',
    displayName: 'Notify team on form warning',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the global CFC team inbox for form submission warnings.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_TEAM_ON_UNMATCHED_FORM_SUBMISSION',
    displayName: 'Notify team on unmatched form submission',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the global CFC team inbox when a form submission cannot be matched to a tracker row.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_TEAM_ON_CLASSIFICATION_NEEDS_REVIEW',
    displayName: 'Notify team when classification needs review',
    defaultValue: false,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the global CFC team inbox when a document cannot be confidently classified.'
  },
  {
    category: 'Notifications',
    key: 'NOTIFY_TEAM_ON_FLASH_REPORT_READY',
    displayName: 'Notify team when flash report ready',
    defaultValue: true,
    valueType: 'boolean',
    allowedValues: 'TRUE, FALSE',
    description: 'Email the global CFC team inbox when a draft flash report is generated.'
  },
  {
    category: 'Localization',
    key: 'TIMEZONE',
    displayName: 'Timezone',
    defaultValue: CONFIG.TIMEZONE,
    valueType: 'string',
    allowedValues: '',
    description: 'Timezone used for dates, deadlines, and trigger scheduling.'
  }
];

let MASTER_CONFIG_CACHE_ = null;
let CONFIG_WARNING_CACHE_ = {};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CFC Reporting')
    .addItem('Run setup', 'setupTracker')
    .addItem('Refresh master config', 'refreshMasterConfigSheet')
    .addItem('Refresh tracker columns/formulas', 'refreshSubmissionTrackerLayout')
    .addItem('Create/update upload form', 'setupUploadForm')
    .addItem('Apply folder sharing settings', 'applyFolderSharingSettings')
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

function onEdit(e) {
  handleMasterConfigEdit_(e);
  handleSubmissionTrackerEdit_(e);
}

function installTriggers() {
  ensureTimeTrigger_('dailyReminderSweep', 'daily', function () {
    return ScriptApp.newTrigger('dailyReminderSweep')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .inTimezone(getConfigString_('TIMEZONE'));
  });

  ensureTimeTrigger_('monthlyRolloverJob', 'monthly', function () {
    return ScriptApp.newTrigger('monthlyRolloverJob')
      .timeBased()
      .onMonthDay(1)
      .atHour(6)
      .inTimezone(getConfigString_('TIMEZONE'));
  });

  replaceTimeTrigger_('driveFolderWatcher', 'watcher', function () {
    return ScriptApp.newTrigger('driveFolderWatcher')
      .timeBased()
      .everyMinutes(getConfigInteger_('DRIVE_WATCHER_INTERVAL_MINUTES'));
  });

  replaceTimeTrigger_('gmailInboxWatcher', 'email watcher', function () {
    return ScriptApp.newTrigger('gmailInboxWatcher')
      .timeBased()
      .everyMinutes(getConfigInteger_('GMAIL_WATCHER_INTERVAL_MINUTES'));
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
    runtimeConfig: getRuntimeConfigSnapshot_(),
    properties: {
      rootFolderId: getScriptProperty_(PROPERTY_KEYS.ROOT_FOLDER_ID),
      trackerSpreadsheetId: getScriptProperty_(PROPERTY_KEYS.TRACKER_SPREADSHEET_ID),
      uploadFormId: getScriptProperty_(PROPERTY_KEYS.UPLOAD_FORM_ID),
      flashTemplateId: getScriptProperty_(PROPERTY_KEYS.FLASH_TEMPLATE_ID)
    }
  };
}
