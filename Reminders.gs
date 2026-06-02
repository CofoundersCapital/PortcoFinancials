function dailyReminderSweep() {
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  const records = getRecords_(sheet);
  const now = new Date();
  let remindersSent = 0;
  let escalationsSent = 0;

  records.forEach(function (entry) {
    const record = entry.record;
    if (!record.company_name || isRecordComplete_(record, docs)) {
      return;
    }

    const daysOverdue = calculateDaysOverdue_(record.deadline, now);
    if (daysOverdue < 1) {
      return;
    }

    const reminderCount = Number(record.reminder_count || 0);
    const lastReminderAt = record.last_reminder_at;

    if (shouldSendReminder_(reminderCount, lastReminderAt, now)) {
      sendReminderEmail_(sheet, entry.rowNumber, record, docs, reminderCount + 1);
      remindersSent++;
      record.reminder_count = reminderCount + 1;
      record.last_reminder_at = now;
    }

    const escalatedAt = record.escalated_at;
    if (shouldEscalate_(daysOverdue, escalatedAt, now)) {
      sendEscalationEmail_(sheet, entry.rowNumber, record, docs, daysOverdue);
      escalationsSent++;
    }
  });

  logEvent_('daily_reminder_sweep', '', '', 'Daily reminder sweep complete', {
    remindersSent: remindersSent,
    escalationsSent: escalationsSent
  });
}

function sendReminderEmail_(sheet, rowNumber, record, docs, reminderNumber) {
  const context = buildEmailContext_(record, docs);
  context.n = reminderNumber;
  context.reminder_count = reminderNumber;

  const template = reminderNumber === 1 ? EMAIL_TEMPLATES.CEO_FIRST_REMINDER : EMAIL_TEMPLATES.CEO_FOLLOW_UP;
  const sent = sendTemplatedEmail_(record.ceo_email, template, context);
  const now = new Date();

  setCellByHeader_(sheet, rowNumber, 'reminder_count', reminderNumber);
  setCellByHeader_(sheet, rowNumber, 'last_reminder_at', now);

  logEvent_('reminder_sent', record.company_name, record.month, 'Reminder sent to CEO', {
    to: sent.to,
    subject: sent.subject,
    reminderNumber: reminderNumber
  });
}

function sendEscalationEmail_(sheet, rowNumber, record, docs, daysOverdue) {
  const context = buildEmailContext_(record, docs);
  context.days_overdue = daysOverdue;

  const sent = sendTemplatedEmail_(record.board_member_email, EMAIL_TEMPLATES.BOARD_ESCALATION, context, {
    cc: record.ceo_email
  });

  setCellByHeader_(sheet, rowNumber, 'escalated_at', new Date());

  logEvent_('escalation_sent', record.company_name, record.month, 'Escalation sent to board member', {
    to: sent.to,
    cc: record.ceo_email,
    subject: sent.subject,
    daysOverdue: daysOverdue
  });
}

function shouldSendReminder_(reminderCount, lastReminderAt, now) {
  if (reminderCount === 0) {
    return true;
  }
  if (!lastReminderAt) {
    return true;
  }
  return daysBetween_(lastReminderAt, now) >= CONFIG.REMINDER_INTERVAL_DAYS;
}

function shouldEscalate_(daysOverdue, escalatedAt, now) {
  if (daysOverdue < CONFIG.ESCALATION_THRESHOLD_DAYS) {
    return false;
  }
  if (!escalatedAt) {
    return true;
  }
  return daysOverdue >= CONFIG.ESCALATION_THRESHOLD_DAYS + CONFIG.REESCALATION_INTERVAL_DAYS
    && daysBetween_(escalatedAt, now) >= CONFIG.REESCALATION_INTERVAL_DAYS;
}

function daysBetween_(fromDate, toDate) {
  const from = startOfDay_(asDate_(fromDate));
  const to = startOfDay_(asDate_(toDate));
  return Math.floor((to.getTime() - from.getTime()) / 86400000);
}
