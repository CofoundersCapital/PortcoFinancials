function onFormSubmit(e) {
  const submission = parseFormSubmission_(e);
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  const rowNumber = findSubmissionRow_(submission.companyName, submission.month);

  if (rowNumber < 0) {
    logEvent_('form_submit_unmatched', submission.companyName, submission.month, 'No tracker row found for form submission', submission);
    GmailApp.sendEmail(
      CONFIG.CFC_TEAM_EMAIL,
      '[CFC] Unmatched monthly financial upload',
      'A form submission could not be matched to a tracker row.\n\n' + safeJsonStringify_(submission)
    );
    return;
  }

  const beforeRecord = getRecordAtRow_(sheet, rowNumber);
  const wasComplete = isRecordComplete_(beforeRecord, docs);
  const targetFolder = getCompanyMonthFolder_(submission.companyName, submission.month, true);

  if (submission.respondentEmail && beforeRecord.ceo_email && normalizeText_(submission.respondentEmail) !== normalizeText_(beforeRecord.ceo_email)) {
    const warning = 'Submitting email ' + submission.respondentEmail + ' does not match CEO email ' + beforeRecord.ceo_email + '.';
    appendNote_(sheet, rowNumber, warning);
    sendCfcWarningEmail_(beforeRecord, warning);
  }

  docs.forEach(function (doc) {
    const fileIds = submission.filesByDocKey[doc.doc_key] || [];
    if (fileIds.length === 0) {
      return;
    }

    const processedCount = fileIds.filter(function (fileId) {
      return processUploadedFile_(fileId, targetFolder, doc);
    }).length;

    if (processedCount > 0) {
      setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_status', 'received');
      setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_date', new Date());
    }
  });

  processSupportingFiles_(submission.supportingFileIds, targetFolder);

  if (submission.notes) {
    appendNote_(sheet, rowNumber, 'CEO note: ' + submission.notes);
  }

  applySubmissionFormulas_(sheet, rowNumber, docs);
  SpreadsheetApp.flush();

  const afterRecord = getRecordAtRow_(sheet, rowNumber);
  if (!wasComplete && isRecordComplete_(afterRecord, docs)) {
    sendCompletionEmail_(afterRecord);
  }

  logEvent_('form_submit_processed', submission.companyName, submission.month, 'Form submission processed', {
    respondentEmail: submission.respondentEmail,
    uploadedDocKeys: Object.keys(submission.filesByDocKey)
  });
}

function driveFolderWatcher() {
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  const records = getRecords_(sheet);
  let updatedRows = 0;

  records.forEach(function (entry) {
    const record = entry.record;
    if (!record.company_name || isRecordComplete_(record, docs)) {
      return;
    }

    const folder = getCompanyMonthFolder_(record.company_name, record.month, false);
    if (!folder) {
      return;
    }

    const wasComplete = isRecordComplete_(record, docs);
    const missingDocs = getMissingDocs_(record, docs);
    const updatedDocKeys = {};
    let rowUpdated = false;

    listFilesInFolder_(folder).forEach(function (file) {
      const matchedDocs = classifyDriveFileDocs_(file, docs, missingDocs).filter(function (doc) {
        return !updatedDocKeys[doc.doc_key];
      });
      if (matchedDocs.length === 0) {
        return;
      }

      saveDriveFileForDocs_(file, folder, matchedDocs);
      matchedDocs.forEach(function (doc) {
        setCellByHeader_(sheet, entry.rowNumber, 'doc_' + doc.doc_key + '_status', 'received');
        setCellByHeader_(sheet, entry.rowNumber, 'doc_' + doc.doc_key + '_date', new Date());
        updatedDocKeys[doc.doc_key] = true;
      });
      rowUpdated = true;
    });

    if (rowUpdated) {
      applySubmissionFormulas_(sheet, entry.rowNumber, docs);
      updatedRows++;
      const afterRecord = getRecordAtRow_(sheet, entry.rowNumber);
      if (!wasComplete && isRecordComplete_(afterRecord, docs)) {
        sendCompletionEmail_(afterRecord);
      }
      logEvent_('drive_watcher_updated_row', record.company_name, record.month, 'Drive watcher marked one or more documents received', {});
    }
  });

  logEvent_('drive_watcher_complete', '', '', 'Drive folder watcher complete', { updatedRows: updatedRows });
}

function gmailInboxWatcher() {
  const sheet = getTrackerSheet_();
  const docs = getRequiredDocs_();
  const processedLabel = getOrCreateGmailLabel_(CONFIG.GMAIL_PROCESSED_LABEL);
  const needsReviewLabel = getOrCreateGmailLabel_(CONFIG.GMAIL_NEEDS_REVIEW_LABEL);
  const query = buildGmailIntakeQuery_();
  const threads = GmailApp.search(query, 0, CONFIG.GMAIL_SEARCH_BATCH_SIZE);
  let processedThreads = 0;
  let reviewThreads = 0;
  let attachmentsSaved = 0;

  threads.forEach(function (thread) {
    const result = processGmailThread_(thread, sheet, docs);

    if (result.attachmentsSaved > 0 && result.needsReview === 0) {
      thread.addLabel(processedLabel);
      processedThreads++;
    } else {
      thread.addLabel(needsReviewLabel);
      reviewThreads++;
    }

    attachmentsSaved += result.attachmentsSaved;
  });

  logEvent_('gmail_inbox_watcher_complete', '', '', 'Gmail inbox watcher complete', {
    query: query,
    threadCount: threads.length,
    processedThreads: processedThreads,
    reviewThreads: reviewThreads,
    attachmentsSaved: attachmentsSaved
  });
}

function processGmailThread_(thread, sheet, docs) {
  const messages = thread.getMessages();
  const result = {
    attachmentsSaved: 0,
    needsReview: 0
  };

  messages.forEach(function (message) {
    const attachments = getMessageAttachments_(message);
    if (attachments.length === 0) {
      return;
    }

    const match = matchSubmissionForEmail_(message);
    if (!match) {
      saveEmailAttachmentsForReview_(attachments, message, 'Could not match email to a tracker company/month.');
      result.needsReview++;
      return;
    }

    const rowNumber = match.rowNumber;
    const record = getRecordAtRow_(sheet, rowNumber);
    const wasComplete = isRecordComplete_(record, docs);
    const folder = getCompanyMonthFolder_(record.company_name, record.month, true);
    const missingDocs = getMissingDocs_(record, docs);
    const updatedDocKeys = {};
    let rowUpdated = false;
    const unmatchedAttachmentNames = [];
    const processedDocKeys = [];

    attachments.forEach(function (attachment) {
      const matchedDocs = classifyEmailAttachmentDocs_(attachment, docs, missingDocs).filter(function (doc) {
        return !updatedDocKeys[doc.doc_key];
      });
      if (matchedDocs.length === 0) {
        unmatchedAttachmentNames.push(attachment.getName());
        saveSingleEmailAttachmentForReview_(attachment, message, 'Could not classify attachment as a required document.');
        return;
      }

      matchedDocs.forEach(function (doc) {
        saveEmailAttachmentToFolder_(attachment, folder, doc);
        setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_status', 'received');
        setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_date', new Date());
        updatedDocKeys[doc.doc_key] = true;
        processedDocKeys.push(doc.doc_key);
        result.attachmentsSaved++;
      });
      rowUpdated = true;
    });

    if (rowUpdated) {
      applySubmissionFormulas_(sheet, rowNumber, docs);
      SpreadsheetApp.flush();
      const afterRecord = getRecordAtRow_(sheet, rowNumber);
      appendNote_(sheet, rowNumber, 'Processed emailed attachments from ' + getEmailAddressFromHeader_(message.getFrom()) + ' for docs [' + uniqueValues_(processedDocKeys).join(', ') + ']: ' + message.getSubject());
      if (!wasComplete && isRecordComplete_(afterRecord, docs)) {
        sendCompletionEmail_(afterRecord);
      }
    }

    if (unmatchedAttachmentNames.length > 0) {
      appendNote_(sheet, rowNumber, 'Email attachments need review: ' + unmatchedAttachmentNames.join(', '));
      result.needsReview++;
    }

    logEvent_('gmail_email_processed', record.company_name, record.month, 'Processed email attachments', {
      from: message.getFrom(),
      subject: message.getSubject(),
      attachmentsSaved: result.attachmentsSaved,
      processedDocKeys: uniqueValues_(processedDocKeys),
      unmatchedAttachmentNames: unmatchedAttachmentNames
    });
  });

  return result;
}

function parseFormSubmission_(e) {
  if (!e) {
    throw new Error('onFormSubmit requires a form submit event.');
  }

  const submission = {
    companyName: '',
    month: '',
    respondentEmail: '',
    notes: '',
    filesByDocKey: {},
    supportingFileIds: []
  };

  if (e.response) {
    submission.respondentEmail = e.response.getRespondentEmail() || '';
    const responses = e.response.getItemResponses();
    responses.forEach(function (itemResponse) {
      const title = itemResponse.getItem().getTitle();
      const response = itemResponse.getResponse();
      applyFormResponseValue_(submission, title, response);
    });
    return validateParsedSubmission_(submission);
  }

  const namedValues = e.namedValues || {};
  Object.keys(namedValues).forEach(function (title) {
    applyFormResponseValue_(submission, title, namedValues[title]);
  });

  return validateParsedSubmission_(submission);
}

function applyFormResponseValue_(submission, title, response) {
  if (title === 'Company') {
    submission.companyName = firstResponseValue_(response);
    return;
  }

  if (title === 'Reporting Month (YYYY-MM)') {
    submission.month = firstResponseValue_(response);
    return;
  }

  if (title === 'Notes') {
    submission.notes = firstResponseValue_(response);
    return;
  }

  if (title === 'Other / Supporting Documents') {
    submission.supportingFileIds = extractDriveIds_(response);
    return;
  }

  const doc = findDocByDisplayName_(title);
  if (doc) {
    submission.filesByDocKey[doc.doc_key] = extractDriveIds_(response);
  }
}

function validateParsedSubmission_(submission) {
  submission.companyName = String(submission.companyName || '').trim();
  submission.month = String(submission.month || '').trim();
  if (!submission.companyName) {
    throw new Error('Form submission is missing Company.');
  }
  if (!isValidMonth_(submission.month)) {
    throw new Error('Form submission Reporting Month must use YYYY-MM.');
  }
  return submission;
}

function firstResponseValue_(response) {
  if (Array.isArray(response)) {
    return String(response[0] || '').trim();
  }
  return String(response || '').trim();
}

function extractDriveIds_(response) {
  const values = Array.isArray(response) ? response : [response];
  const ids = [];
  values.forEach(function (value) {
    if (Array.isArray(value)) {
      value.forEach(function (nested) {
        ids.push.apply(ids, extractDriveIds_(nested));
      });
      return;
    }
    const text = String(value || '').trim();
    if (!text) {
      return;
    }
    const urlMatches = text.match(/[-\w]{25,}/g);
    if (urlMatches) {
      urlMatches.forEach(function (id) {
        ids.push(id);
      });
    }
  });
  return ids.filter(function (id, index) {
    return ids.indexOf(id) === index;
  });
}

function processUploadedFile_(fileId, targetFolder, doc) {
  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (err) {
    logEvent_('upload_file_link_error', '', '', 'Could not open pasted Drive file link', {
      fileId: fileId,
      docKey: doc.doc_key,
      error: err.message
    });
    return false;
  }

  const extension = getFileExtension_(file.getName()) || extensionFromMimeType_(file.getMimeType()) || splitExtensions_(doc.accepted_extensions)[0] || 'file';
  const accepted = splitExtensions_(doc.accepted_extensions);

  if (accepted.length > 0 && accepted.indexOf(extension) === -1) {
    logEvent_('upload_extension_warning', '', '', 'Uploaded file extension did not match config', {
      fileName: file.getName(),
      extension: extension,
      accepted: accepted,
      docKey: doc.doc_key
    });
  }

  trashExistingFileNamed_(targetFolder, doc.doc_key + '.' + extension, file.getId());
  file.setName(doc.doc_key + '.' + extension);
  moveFileToFolder_(file, targetFolder);
  return true;
}

function processSupportingFiles_(fileIds, targetFolder) {
  (fileIds || []).forEach(function (fileId, index) {
    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (err) {
      logEvent_('supporting_file_link_error', '', '', 'Could not open pasted supporting Drive file link', {
        fileId: fileId,
        error: err.message
      });
      return;
    }

    const extension = getFileExtension_(file.getName());
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss');
    const name = 'supporting-' + timestamp + '-' + (index + 1) + (extension ? '.' + extension : '');
    trashExistingFileNamed_(targetFolder, name, file.getId());
    file.setName(name);
    moveFileToFolder_(file, targetFolder);
  });
}

function listFilesInFolder_(folder) {
  const files = [];
  const iterator = folder.getFiles();
  while (iterator.hasNext()) {
    files.push(iterator.next());
  }
  return files;
}

function findMatchingDocumentFile_(folder, doc) {
  const files = folder.getFiles();
  const acceptedExtensions = splitExtensions_(doc.accepted_extensions);
  const keyPattern = buildDocFilenamePattern_(doc);

  while (files.hasNext()) {
    const file = files.next();
    const extension = getFileExtension_(file.getName()) || extensionFromMimeType_(file.getMimeType());
    if (acceptedExtensions.length > 0 && acceptedExtensions.indexOf(extension) === -1) {
      continue;
    }
    if (keyPattern.test(file.getName())) {
      return file;
    }
  }

  return null;
}

function classifyDriveFileDocs_(file, docs, missingDocs) {
  return classifyFileForDocs_(file.getName(), file.getMimeType(), docs, missingDocs);
}

function saveDriveFileForDocs_(file, folder, matchedDocs) {
  if (matchedDocs.length === 1) {
    renameToCanonical_(file, matchedDocs[0]);
    return;
  }

  matchedDocs.forEach(function (doc) {
    copyDriveFileToCanonical_(file, folder, doc);
  });
}

function copyDriveFileToCanonical_(file, folder, doc) {
  const extension = getFileExtension_(file.getName()) || extensionFromMimeType_(file.getMimeType()) || splitExtensions_(doc.accepted_extensions)[0] || 'file';
  const canonicalName = doc.doc_key + '.' + extension;
  if (file.getName() === canonicalName) {
    return file;
  }

  trashExistingFileNamed_(folder, canonicalName, file.getId());
  return file.makeCopy(canonicalName, folder);
}

function buildDocFilenamePattern_(doc) {
  let terms = [doc.doc_key].concat(String(doc.display_name || '').split(/[^a-zA-Z0-9]+/));
  if (doc.doc_key === 'financials') {
    terms = terms.concat(['financial', 'financials', 'statement', 'income', 'pnl', 'p&l']);
  }
  if (doc.doc_key === 'model') {
    terms = terms.concat(['model']);
  }
  if (doc.doc_key === 'forecast') {
    terms = terms.concat(['forecast', 'budget']);
  }
  const unique = terms
    .map(function (term) {
      return String(term || '').toLowerCase();
    })
    .filter(function (term) {
      return term.length >= 3;
    })
    .filter(function (term, index, list) {
      return list.indexOf(term) === index;
    });
  return new RegExp('(' + unique.map(escapeRegExp_).join('|') + ')', 'i');
}

function renameToCanonical_(file, doc) {
  const extension = getFileExtension_(file.getName()) || extensionFromMimeType_(file.getMimeType());
  if (!extension) {
    return;
  }
  const canonicalName = doc.doc_key + '.' + extension;
  if (file.getName() === canonicalName) {
    return;
  }
  const parents = file.getParents();
  if (parents.hasNext()) {
    trashExistingFileNamed_(parents.next(), canonicalName, file.getId());
  }
  file.setName(canonicalName);
}

function trashExistingFileNamed_(folder, name, exceptFileId) {
  const files = folder.getFilesByName(name);
  while (files.hasNext()) {
    const existing = files.next();
    if (existing.getId() !== exceptFileId) {
      existing.setTrashed(true);
    }
  }
}

function extensionFromMimeType_(mimeType) {
  const map = {};
  map[MimeType.PDF] = 'pdf';
  map[MimeType.MICROSOFT_EXCEL] = 'xlsx';
  map[MimeType.GOOGLE_SHEETS] = 'xlsx';
  map[MimeType.CSV] = 'csv';
  map[MimeType.PLAIN_TEXT] = 'txt';
  return map[mimeType] || '';
}

function setupEmailIntake_() {
  getOrCreateGmailLabel_(CONFIG.GMAIL_PROCESSED_LABEL);
  getOrCreateGmailLabel_(CONFIG.GMAIL_NEEDS_REVIEW_LABEL);
  logEvent_('email_intake_setup', '', '', 'Email intake labels checked', {
    intakeEmail: CONFIG.INTAKE_EMAIL,
    processedLabel: CONFIG.GMAIL_PROCESSED_LABEL,
    needsReviewLabel: CONFIG.GMAIL_NEEDS_REVIEW_LABEL
  });
}

function buildGmailIntakeQuery_() {
  const parts = ['has:attachment'];
  if (CONFIG.INTAKE_EMAIL) {
    parts.push('to:' + CONFIG.INTAKE_EMAIL);
  }
  if (CONFIG.GMAIL_LOOKBACK_DAYS) {
    parts.push('newer_than:' + CONFIG.GMAIL_LOOKBACK_DAYS + 'd');
  }
  parts.push('-label:' + CONFIG.GMAIL_PROCESSED_LABEL);
  parts.push('-label:' + CONFIG.GMAIL_NEEDS_REVIEW_LABEL);
  return parts.join(' ');
}

function getOrCreateGmailLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function getMessageAttachments_(message) {
  try {
    return message.getAttachments({
      includeInlineImages: false,
      includeAttachments: true
    }).filter(isLikelyFinancialAttachment_);
  } catch (err) {
    return message.getAttachments().filter(isLikelyFinancialAttachment_);
  }
}

function isLikelyFinancialAttachment_(attachment) {
  const name = attachment.getName();
  const extension = getFileExtension_(name);
  const contentType = String(attachment.getContentType() || '').toLowerCase();
  if (!name || /^image\//.test(contentType)) {
    return false;
  }
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'heic', 'ics', 'vcf'].indexOf(extension) >= 0) {
    return false;
  }
  return true;
}

function matchSubmissionForEmail_(message) {
  const records = getRecords_(getTrackerSheet_()).filter(function (entry) {
    return entry.record.company_name;
  });
  const fromEmail = getEmailAddressFromHeader_(message.getFrom());
  const subject = message.getSubject() || '';
  const body = truncate_(message.getPlainBody() || '', 20000);
  const searchableText = [subject, body].join('\n');
  const inferredMonth = inferReportingMonthFromText_(searchableText, message.getDate());
  const companyName = inferCompanyNameFromEmail_(records, fromEmail, searchableText);

  if (!companyName) {
    return null;
  }

  const companyRecords = records.filter(function (entry) {
    return normalizeText_(entry.record.company_name) === normalizeText_(companyName);
  });

  if (inferredMonth) {
    const exact = companyRecords.filter(function (entry) {
      return String(entry.record.month) === inferredMonth;
    })[0];
    if (exact) {
      return exact;
    }
  }

  return chooseBestSubmissionRow_(companyRecords);
}

function inferCompanyNameFromEmail_(records, fromEmail, searchableText) {
  const senderMatches = records.filter(function (entry) {
    return normalizeText_(entry.record.ceo_email) === normalizeText_(fromEmail);
  });

  const uniqueSenderCompanies = uniqueValues_(senderMatches.map(function (entry) {
    return entry.record.company_name;
  }));

  if (uniqueSenderCompanies.length === 1) {
    return uniqueSenderCompanies[0];
  }

  const domainMatches = inferCompanyNamesBySenderDomain_(records, fromEmail);
  if (domainMatches.length === 1) {
    return domainMatches[0];
  }

  const textMatches = uniqueValues_(records
    .map(function (entry) {
      return entry.record.company_name;
    })
    .filter(function (companyName) {
      return containsCompanyName_(searchableText, companyName);
    }));

  if (textMatches.length === 1) {
    return textMatches[0];
  }

  if (uniqueSenderCompanies.length > 1) {
    const senderTextMatches = uniqueSenderCompanies.filter(function (companyName) {
      return containsCompanyName_(searchableText, companyName);
    });
    if (senderTextMatches.length === 1) {
      return senderTextMatches[0];
    }
  }

  return '';
}

function inferCompanyNamesBySenderDomain_(records, fromEmail) {
  const senderDomain = getEmailDomain_(fromEmail);
  if (!senderDomain || isGenericEmailDomain_(senderDomain)) {
    return [];
  }

  return uniqueValues_(records
    .filter(function (entry) {
      return getEmailDomain_(entry.record.ceo_email) === senderDomain;
    })
    .map(function (entry) {
      return entry.record.company_name;
    }));
}

function chooseBestSubmissionRow_(companyRecords) {
  if (companyRecords.length === 0) {
    return null;
  }

  const docs = getRequiredDocs_();
  const incomplete = companyRecords.filter(function (entry) {
    return !isRecordComplete_(entry.record, docs);
  });
  const pool = incomplete.length ? incomplete : companyRecords;
  pool.sort(function (a, b) {
    return String(b.record.month).localeCompare(String(a.record.month));
  });
  return pool[0];
}

function classifyEmailAttachmentDocs_(attachment, docs, missingDocs) {
  return classifyFileForDocs_(attachment.getName(), attachment.getContentType(), docs, missingDocs);
}

function classifyFileForDocs_(fileName, contentType, docs, missingDocs) {
  const extension = getFileExtension_(fileName) || extensionFromMimeType_(contentType);
  const candidateDocs = docs.filter(function (doc) {
    const accepted = splitExtensions_(doc.accepted_extensions);
    return accepted.length === 0 || accepted.indexOf(extension) >= 0;
  });

  if (candidateDocs.length === 0) {
    return [];
  }

  const scored = candidateDocs.map(function (doc) {
    return {
      doc: doc,
      score: scoreDocFileMatch_(doc, fileName, missingDocs)
    };
  }).sort(function (a, b) {
    return b.score - a.score;
  });

  const strongMatches = scored.filter(function (entry) {
    return entry.score >= 10;
  });

  if (strongMatches.length > 0) {
    return strongMatches.map(function (entry) {
      return entry.doc;
    });
  }

  const missingCandidates = candidateDocs.filter(function (doc) {
    return isDocStillMissing_(doc, missingDocs);
  });
  if (missingCandidates.length === 1) {
    return [missingCandidates[0]];
  }

  return [];
}

function scoreDocFileMatch_(doc, fileName, missingDocs) {
  let score = 0;
  if (hasStrongDocKeyword_(doc, fileName)) {
    score += 10;
  }
  if (isDocStillMissing_(doc, missingDocs)) {
    score += 2;
  }
  return score;
}

function hasStrongDocKeyword_(doc, fileName) {
  const name = String(fileName || '').toLowerCase();
  if (doc.doc_key === 'financials') {
    return /\b(financials|monthly financials|financial statements?|income statements?|p&l|pnl|balance sheets?|cash flows?|statement of cash flows?)\b/i.test(name);
  }
  if (doc.doc_key === 'model') {
    return /\b(model|operating model|financial model)\b/i.test(name);
  }
  if (doc.doc_key === 'forecast') {
    return /\b(forecast|budget|plan|projection|projections)\b/i.test(name);
  }
  return buildDocFilenamePattern_(doc).test(fileName);
}

function saveEmailAttachmentToFolder_(attachment, folder, doc) {
  const extension = getFileExtension_(attachment.getName()) || extensionFromMimeType_(attachment.getContentType()) || splitExtensions_(doc.accepted_extensions)[0] || 'file';
  const canonicalName = doc.doc_key + '.' + extension;
  trashExistingFileNamed_(folder, canonicalName, '');
  const file = folder.createFile(attachment.copyBlob()).setName(canonicalName);
  return file;
}

function saveEmailAttachmentsForReview_(attachments, message, reason) {
  attachments.forEach(function (attachment) {
    saveSingleEmailAttachmentForReview_(attachment, message, reason);
  });
}

function saveSingleEmailAttachmentForReview_(attachment, message, reason) {
  const folder = getEmailNeedsReviewFolder_(true);
  const dateStamp = Utilities.formatDate(message.getDate(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss');
  const safeSubject = String(message.getSubject() || 'no-subject').replace(/[^a-zA-Z0-9._-]+/g, '-').slice(0, 80);
  const fileName = dateStamp + '-' + safeSubject + '-' + attachment.getName();
  folder.createFile(attachment.copyBlob()).setName(fileName);
  logEvent_('gmail_attachment_needs_review', '', '', reason, {
    from: message.getFrom(),
    subject: message.getSubject(),
    attachmentName: attachment.getName(),
    savedAs: fileName
  });
}

function getEmailNeedsReviewFolder_(createIfMissing) {
  const rootFolder = getRootFolder_(createIfMissing);
  if (!rootFolder) {
    return null;
  }
  return getNamedFolder_(rootFolder, '_email_needs_review', createIfMissing);
}

function inferReportingMonthFromText_(text, messageDate) {
  const value = String(text || '');
  const numeric = value.match(/\b(20\d{2})[-_/](0[1-9]|1[0-2])\b/);
  if (numeric) {
    return numeric[1] + '-' + numeric[2];
  }

  const monthPattern = /\b(january|jan|february|feb|march|mar|april|apr|may|june|jun|july|jul|august|aug|september|sep|sept|october|oct|november|nov|december|dec)\s*(20\d{2})?\b/i;
  const monthMatch = value.match(monthPattern);
  if (!monthMatch) {
    return '';
  }

  const monthNumber = monthNameToNumber_(monthMatch[1]);
  if (!monthNumber) {
    return '';
  }

  const emailDate = messageDate || new Date();
  let year = monthMatch[2] ? Number(monthMatch[2]) : Number(Utilities.formatDate(emailDate, CONFIG.TIMEZONE, 'yyyy'));
  const emailMonth = Number(Utilities.formatDate(emailDate, CONFIG.TIMEZONE, 'M'));
  if (!monthMatch[2] && monthNumber > emailMonth + 1) {
    year--;
  }
  return year + '-' + String(monthNumber).padStart(2, '0');
}

function monthNameToNumber_(monthName) {
  const key = String(monthName || '').toLowerCase().slice(0, 3);
  const map = {
    jan: 1,
    feb: 2,
    mar: 3,
    apr: 4,
    may: 5,
    jun: 6,
    jul: 7,
    aug: 8,
    sep: 9,
    oct: 10,
    nov: 11,
    dec: 12
  };
  return map[key] || 0;
}

function getEmailAddressFromHeader_(header) {
  const text = String(header || '').trim();
  const angleMatch = text.match(/<([^>]+)>/);
  if (angleMatch) {
    return angleMatch[1].trim().toLowerCase();
  }
  const emailMatch = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return emailMatch ? emailMatch[0].toLowerCase() : text.toLowerCase();
}

function getEmailDomain_(email) {
  const address = getEmailAddressFromHeader_(email);
  const parts = address.split('@');
  return parts.length === 2 ? parts[1].toLowerCase() : '';
}

function isGenericEmailDomain_(domain) {
  return [
    'gmail.com',
    'googlemail.com',
    'outlook.com',
    'hotmail.com',
    'icloud.com',
    'me.com',
    'yahoo.com',
    'aol.com',
    'proton.me',
    'protonmail.com'
  ].indexOf(String(domain || '').toLowerCase()) >= 0;
}

function containsCompanyName_(text, companyName) {
  const normalizedText = normalizeText_(text);
  const normalizedCompany = normalizeText_(companyName);
  if (!normalizedCompany) {
    return false;
  }
  return normalizedText.indexOf(normalizedCompany) >= 0;
}

function isDocStillMissing_(doc, missingDocs) {
  return missingDocs.some(function (missingDoc) {
    return missingDoc.doc_key === doc.doc_key;
  });
}

function uniqueValues_(values) {
  return values.filter(function (value, index, list) {
    return value && list.indexOf(value) === index;
  });
}
