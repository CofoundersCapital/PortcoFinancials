const EMAIL_TEMPLATES = {
  CEO_FIRST_REMINDER: {
    subject: '[CFC] Monthly financials for {{month}} - friendly reminder',
    body: [
      'Hi {{ceo_name}},',
      '',
      'This is a quick reminder that your {{month}} financial materials for {{company_name}} were due on {{deadline}}. We are still missing:',
      '',
      '{{missing_docs_bulleted}}',
      '',
      'Please reply with the missing files attached, or email them to {{intake_email}} with {{company_name}} and {{month}} in the subject line.',
      '',
      'You can also drop files in your shared folder: {{folder_url}}',
      '',
      'Thanks,',
      'CFC Team (automated reminder)'
    ].join('\n')
  },
  CEO_FOLLOW_UP: {
    subject: '[CFC][Follow-up #{{n}}] Monthly financials for {{month}}',
    body: [
      'Hi {{ceo_name}},',
      '',
      'Following up on the {{month}} financial materials for {{company_name}}. We are still missing:',
      '',
      '{{missing_docs_bulleted}}',
      '',
      'Please reply with the missing files attached, or email them to {{intake_email}} with {{company_name}} and {{month}} in the subject line.',
      '',
      'You can also drop files in your shared folder: {{folder_url}}',
      '',
      'Thanks,',
      'CFC Team (automated reminder)'
    ].join('\n')
  },
  BOARD_ESCALATION: {
    subject: '[CFC][Escalation] {{company_name}} - {{days_overdue}} days overdue',
    body: [
      'Hi {{board_member_name}},',
      '',
      '{{company_name}} is {{days_overdue}} days past the {{month}} reporting deadline.',
      '',
      'Missing documents:',
      '{{missing_docs_bulleted}}',
      '',
      'CEO contact: {{ceo_name}} <{{ceo_email}}>',
      'Reminders sent: {{reminder_count}} (most recent: {{last_reminder_at}})',
      '',
      'Company folder: {{folder_url}}',
      '',
      '- CFC Reporting System'
    ].join('\n')
  },
  CEO_DOCUMENT_RECEIVED: {
    subject: '[CFC] {{company_name}} {{month}} document received',
    body: [
      'Hi {{ceo_name}},',
      '',
      'We received and processed the following {{month}} material for {{company_name}}:',
      '',
      '{{received_docs_bulleted}}',
      '',
      'Source file: {{file_name}}',
      '',
      '- CFC Team'
    ].join('\n')
  },
  BOARD_DOCUMENT_RECEIVED: {
    subject: '[CFC] {{company_name}} {{month}} document received',
    body: [
      'Hi {{board_member_name}},',
      '',
      'The following {{month}} material was received for {{company_name}}:',
      '',
      '{{received_docs_bulleted}}',
      '',
      'Source: {{source}}',
      'Source file: {{file_name}}',
      'Company folder: {{folder_url}}',
      'Tracker: {{tracker_url}}',
      '',
      '- CFC Reporting System'
    ].join('\n')
  },
  CEO_COMPLETION: {
    subject: '[CFC] {{month}} submission received - thank you',
    body: [
      'Hi {{ceo_name}},',
      '',
      'We have received all required materials for {{company_name}} for {{month}}. No further action needed.',
      '',
      '- CFC Team'
    ].join('\n')
  },
  BOARD_COMPLETION: {
    subject: '[CFC] {{company_name}} {{month}} submission complete',
    body: [
      'Hi {{board_member_name}},',
      '',
      'All required {{month}} reporting materials for {{company_name}} have been received.',
      '',
      'Company folder: {{folder_url}}',
      'Tracker: {{tracker_url}}',
      '',
      '- CFC Reporting System'
    ].join('\n')
  },
  CFC_REVIEW_FLASH_REPORT: {
    subject: '[CFC] Draft flash report ready - {{company_name}} {{month}}',
    body: [
      'A draft flash report is ready for review.',
      '',
      'Company: {{company_name}}',
      'Reporting month: {{month}}',
      'Extraction confidence: {{extraction_confidence}}',
      '',
      'Report: {{flash_report_url}}',
      'Tracker: {{tracker_url}}',
      '',
      'Please review extraction notes and source materials before using this report.'
    ].join('\n')
  },
  CFC_FORM_WARNING: {
    subject: '[CFC] Upload warning - {{company_name}} {{month}}',
    body: [
      'An upload warning was recorded.',
      '',
      'Company: {{company_name}}',
      'Reporting month: {{month}}',
      'Warning: {{warning}}',
      '',
      'Tracker: {{tracker_url}}',
      'Company folder: {{folder_url}}'
    ].join('\n')
  },
  CFC_CLASSIFICATION_NEEDS_REVIEW: {
    subject: '[CFC] Document needs review - {{company_name}} {{month}}',
    body: [
      'A document could not be confidently classified.',
      '',
      'Company: {{company_name}}',
      'Reporting month: {{month}}',
      'Source: {{source}}',
      'File: {{file_name}}',
      'File URL: {{file_url}}',
      'Reason: {{reason}}',
      '',
      'Tracker: {{tracker_url}}',
      'Company folder: {{folder_url}}'
    ].join('\n')
  }
};

function renderTemplate_(template, data) {
  let rendered = String(template || '');
  Object.keys(data || {}).forEach(function (key) {
    rendered = rendered.replace(new RegExp('{{' + escapeRegExp_(key) + '}}', 'g'), stringifyTemplateValue_(data[key]));
  });
  return rendered;
}

function stringifyTemplateValue_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return formatDateTime_(value);
  }
  if (Array.isArray(value)) {
    return value.join('\n');
  }
  return String(value);
}

function sendTemplatedEmail_(to, template, data, options) {
  if (!to) {
    throw new Error('Cannot send email without a recipient.');
  }

  const subject = renderTemplate_(template.subject, data);
  const body = renderTemplate_(template.body, data);
  GmailApp.sendEmail(to, subject, body, options || {});
  return {
    to: to,
    subject: subject,
    body: body
  };
}

function buildEmailContext_(record, docs) {
  const companyFolder = getCompanyFolder_(record.company_name, false);
  const missingDocs = getMissingDocs_(record, docs);
  const daysOverdue = calculateDaysOverdue_(record.deadline);

  return {
    month: record.month,
    company_name: record.company_name,
    ceo_name: record.ceo_name,
    ceo_email: record.ceo_email,
    board_member_name: record.board_member_name,
    board_member_email: record.board_member_email,
    deadline: formatDate_(record.deadline),
    days_overdue: daysOverdue,
    reminder_count: Number(record.reminder_count || 0),
    last_reminder_at: record.last_reminder_at ? formatDateTime_(record.last_reminder_at) : 'none',
    missing_docs_bulleted: missingDocs.map(function (doc) {
      return '- ' + doc.display_name;
    }).join('\n'),
    intake_email: getConfigString_('INTAKE_EMAIL'),
    form_url: getFormUrl_(),
    folder_url: companyFolder ? companyFolder.getUrl() : '',
    tracker_url: getTrackerSpreadsheet_().getUrl()
  };
}

function sendCompletionEmail_(record) {
  sendCompletionNotifications_(record);
}

function sendCompletionNotifications_(record) {
  const context = buildEmailContext_(record, getRequiredDocs_());
  if (isFeatureEnabled_('NOTIFY_FOUNDER_ON_SUBMISSION_COMPLETE') && record.ceo_email) {
    const founderSent = sendTemplatedEmail_(record.ceo_email, EMAIL_TEMPLATES.CEO_COMPLETION, context);
    logEvent_('completion_email_sent', record.company_name, record.month, 'Completion confirmation sent to CEO', {
      to: founderSent.to,
      subject: founderSent.subject
    });
  }

  if (isFeatureEnabled_('NOTIFY_BOARD_ON_SUBMISSION_COMPLETE') && record.board_member_email) {
    const boardSent = sendTemplatedEmail_(record.board_member_email, EMAIL_TEMPLATES.BOARD_COMPLETION, context);
    logEvent_('board_completion_email_sent', record.company_name, record.month, 'Completion confirmation sent to board member', {
      to: boardSent.to,
      subject: boardSent.subject
    });
  }
}

function sendDocumentReceivedNotifications_(record, matchedDocs, source, fileName) {
  if (!matchedDocs || matchedDocs.length === 0) {
    return;
  }

  const context = buildEmailContext_(record, getRequiredDocs_());
  context.received_docs_bulleted = matchedDocs.map(function (doc) {
    return '- ' + doc.display_name;
  }).join('\n');
  context.received_docs_csv = matchedDocs.map(function (doc) {
    return doc.display_name;
  }).join(', ');
  context.source = source || '';
  context.file_name = fileName || '';

  if (isFeatureEnabled_('NOTIFY_FOUNDER_ON_DOCUMENT_RECEIVED') && record.ceo_email) {
    const founderSent = sendTemplatedEmail_(record.ceo_email, EMAIL_TEMPLATES.CEO_DOCUMENT_RECEIVED, context);
    logEvent_('founder_document_received_email_sent', record.company_name, record.month, 'Document received confirmation sent to CEO', {
      to: founderSent.to,
      subject: founderSent.subject,
      docKeys: matchedDocs.map(function (doc) { return doc.doc_key; })
    });
  }

  if (isFeatureEnabled_('NOTIFY_BOARD_ON_DOCUMENT_RECEIVED') && record.board_member_email) {
    const boardSent = sendTemplatedEmail_(record.board_member_email, EMAIL_TEMPLATES.BOARD_DOCUMENT_RECEIVED, context);
    logEvent_('board_document_received_email_sent', record.company_name, record.month, 'Document received notification sent to board member', {
      to: boardSent.to,
      subject: boardSent.subject,
      docKeys: matchedDocs.map(function (doc) { return doc.doc_key; })
    });
  }
}

function sendCfcWarningEmail_(record, warning) {
  if (!isFeatureEnabled_('NOTIFY_TEAM_ON_FORM_WARNING')) {
    return;
  }

  const context = buildEmailContext_(record, getRequiredDocs_());
  context.warning = warning;
  sendTemplatedEmail_(getConfigString_('CFC_TEAM_EMAIL'), EMAIL_TEMPLATES.CFC_FORM_WARNING, context);
}

function sendUnmatchedFormSubmissionEmail_(submission) {
  if (!isFeatureEnabled_('NOTIFY_TEAM_ON_UNMATCHED_FORM_SUBMISSION')) {
    return;
  }

  GmailApp.sendEmail(
    getConfigString_('CFC_TEAM_EMAIL'),
    '[CFC] Unmatched monthly financial upload',
    'A form submission could not be matched to a tracker row.\n\n' + safeJsonStringify_(submission)
  );
}

function sendClassificationNeedsReviewEmail_(record, source, fileName, fileUrl, reason) {
  if (!isFeatureEnabled_('NOTIFY_TEAM_ON_CLASSIFICATION_NEEDS_REVIEW')) {
    return;
  }

  const context = buildEmailContext_(record, getRequiredDocs_());
  context.source = source || '';
  context.file_name = fileName || '';
  context.file_url = fileUrl || '';
  context.reason = reason || '';
  sendTemplatedEmail_(getConfigString_('CFC_TEAM_EMAIL'), EMAIL_TEMPLATES.CFC_CLASSIFICATION_NEEDS_REVIEW, context);
}
