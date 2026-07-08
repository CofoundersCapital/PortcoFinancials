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

  const updatedDocKeys = {};
  const processedDocKeys = [];
  const submittedFileIds = getSubmittedChecklistFileIds_(submission);
  submittedFileIds.forEach(function (fileId) {
    const classification = processSubmittedChecklistFile_(fileId, docs, getMissingDocs_(beforeRecord, docs), sheet, rowNumber, beforeRecord);
    const matchedDocs = filterMatchedDocsForUpdate_(classification.matchedDocs, updatedDocKeys);
    if (classification.matchedDocs.length > 0 && matchedDocs.length === 0) {
      return;
    }
    if (matchedDocs.length > 0 && classification.file) {
      saveDriveFileForDocs_(classification.file, targetFolder, matchedDocs, true);
    }
    markMatchedDocsReceived_(sheet, rowNumber, matchedDocs, updatedDocKeys, processedDocKeys);
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
    uploadedDocKeys: uniqueValues_(processedDocKeys),
    submittedFileCount: submittedFileIds.length
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
      const classification = classifyDriveFileForDocs_(file, docs, missingDocs);
      const matchedDocs = filterMatchedDocsForUpdate_(classification.matchedDocs, updatedDocKeys);
      if (classification.matchedDocs.length === 0) {
        recordDocumentClassificationNeedsReview_(sheet, entry.rowNumber, record, 'drive_watcher', file, classification.reason, classification.raw);
        return;
      }
      if (matchedDocs.length === 0) {
        return;
      }

      saveDriveFileForDocs_(file, folder, matchedDocs, false);
      markMatchedDocsReceived_(sheet, entry.rowNumber, matchedDocs, updatedDocKeys, []);
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
      const classification = classifyEmailAttachmentForDocs_(attachment, docs, missingDocs);
      const matchedDocs = filterMatchedDocsForUpdate_(classification.matchedDocs, updatedDocKeys);
      if (classification.matchedDocs.length === 0) {
        unmatchedAttachmentNames.push(attachment.getName());
        saveSingleEmailAttachmentForReview_(attachment, message, classification.reason || 'Could not classify attachment as a required document.');
        return;
      }
      if (matchedDocs.length === 0) {
        return;
      }

      matchedDocs.forEach(function (doc) {
        saveEmailAttachmentToFolder_(attachment, folder, doc);
        result.attachmentsSaved++;
      });
      markMatchedDocsReceived_(sheet, rowNumber, matchedDocs, updatedDocKeys, processedDocKeys);
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

function getSubmittedChecklistFileIds_(submission) {
  const ids = [];
  Object.keys(submission.filesByDocKey || {}).forEach(function (docKey) {
    (submission.filesByDocKey[docKey] || []).forEach(function (fileId) {
      ids.push(fileId);
    });
  });
  return ids.filter(function (id, index) {
    return id && ids.indexOf(id) === index;
  });
}

function processSubmittedChecklistFile_(fileId, docs, missingDocs, sheet, rowNumber, record) {
  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (err) {
    logEvent_('upload_file_link_error', '', '', 'Could not open pasted Drive file link', {
      fileId: fileId,
      error: err.message
    });
    return buildDocumentClassificationResult_([], 'Could not open pasted Drive file link: ' + err.message, null);
  }

  const classification = classifyDriveFileForDocs_(file, docs, missingDocs);
  classification.file = file;
  if (classification.matchedDocs.length === 0) {
    recordDocumentClassificationNeedsReview_(sheet, rowNumber, record, 'form_upload', file, classification.reason, classification.raw);
  }
  return classification;
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

function classifyDriveFileForDocs_(file, docs, missingDocs, overrides) {
  const fileName = overrides && overrides.fileName ? overrides.fileName : file.getName();
  const mimeType = overrides && overrides.mimeType ? overrides.mimeType : file.getMimeType();
  const extension = getFileExtension_(fileName) || extensionFromMimeType_(mimeType);
  const candidateDocs = getCandidateDocsForExtension_(docs, extension);

  if (candidateDocs.length === 0) {
    return buildDocumentClassificationResult_([], 'No required document types accept extension "' + (extension || 'unknown') + '".', null);
  }

  let extractedText = '';
  try {
    extractedText = truncate_(extractFileText_(file), CONFIG.OPENAI_CLASSIFICATION_MAX_CHARS);
  } catch (err) {
    extractedText = '[Text extraction failed: ' + err.message + ']';
  }

  let raw;
  try {
    raw = openaiClassifyDocument_({
      fileName: fileName,
      mimeType: mimeType,
      extension: extension,
      candidateDocs: candidateDocs,
      missingDocs: missingDocs || [],
      extractedText: extractedText
    });
  } catch (err) {
    return buildDocumentClassificationResult_([], 'LLM document classification failed: ' + err.message, null);
  }

  return normalizeDocumentClassification_(raw, candidateDocs);
}

function classifyEmailAttachmentForDocs_(attachment, docs, missingDocs) {
  let tempFile = null;
  try {
    tempFile = createTemporaryClassificationFile_(attachment);
    return classifyDriveFileForDocs_(tempFile, docs, missingDocs, {
      fileName: attachment.getName(),
      mimeType: attachment.getContentType()
    });
  } catch (err) {
    return buildDocumentClassificationResult_([], 'Could not prepare attachment for LLM classification: ' + err.message, null);
  } finally {
    if (tempFile) {
      try {
        tempFile.setTrashed(true);
      } catch (cleanupErr) {
        logEvent_('classification_temp_cleanup_error', '', '', 'Could not trash temporary classification file', {
          fileName: tempFile.getName(),
          error: cleanupErr.message
        });
      }
    }
  }
}

function createTemporaryClassificationFile_(attachment) {
  const folder = getTemporaryClassificationFolder_(true);
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss');
  return folder.createFile(attachment.copyBlob()).setName('classify-' + timestamp + '-' + attachment.getName());
}

function getTemporaryClassificationFolder_(createIfMissing) {
  const rootFolder = getRootFolder_(createIfMissing);
  if (!rootFolder) {
    return null;
  }
  return getNamedFolder_(rootFolder, '_classification_temp', createIfMissing);
}

function getCandidateDocsForExtension_(docs, extension) {
  return docs.filter(function (doc) {
    const accepted = splitExtensions_(doc.accepted_extensions);
    return accepted.length === 0 || accepted.indexOf(extension) >= 0;
  });
}

function normalizeDocumentClassification_(raw, candidateDocs) {
  const docsByKey = {};
  candidateDocs.forEach(function (doc) {
    docsByKey[doc.doc_key] = doc;
  });

  const seen = {};
  const matches = [];
  (raw && raw.matches ? raw.matches : []).forEach(function (match) {
    if (!match
      || !docsByKey[match.doc_key]
      || ['high', 'medium'].indexOf(match.confidence) < 0
      || seen[match.doc_key]) {
      return;
    }

    seen[match.doc_key] = true;
    matches.push(docsByKey[match.doc_key]);
  });

  const reason = matches.length > 0
    ? ''
    : buildDocumentClassificationNoMatchReason_(raw);

  return buildDocumentClassificationResult_(matches, reason, raw);
}

function buildDocumentClassificationNoMatchReason_(raw) {
  if (raw && raw.unmatched_reason) {
    return String(raw.unmatched_reason);
  }

  const lowConfidence = raw && raw.matches ? raw.matches.filter(function (match) {
    return match && match.confidence === 'low';
  }) : [];
  if (lowConfidence.length > 0) {
    return 'LLM returned only low-confidence document matches.';
  }

  return 'LLM returned no high- or medium-confidence required document matches.';
}

function buildDocumentClassificationResult_(matchedDocs, reason, raw) {
  return {
    matchedDocs: matchedDocs || [],
    reason: reason || '',
    raw: raw || null
  };
}

function filterMatchedDocsForUpdate_(matchedDocs, updatedDocKeys) {
  return (matchedDocs || []).filter(function (doc) {
    return !updatedDocKeys[doc.doc_key];
  });
}

function markMatchedDocsReceived_(sheet, rowNumber, matchedDocs, updatedDocKeys, processedDocKeys) {
  (matchedDocs || []).forEach(function (doc) {
    setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_status', 'received');
    setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_date', new Date());
    updatedDocKeys[doc.doc_key] = true;
    if (processedDocKeys) {
      processedDocKeys.push(doc.doc_key);
    }
  });
}

function recordDocumentClassificationNeedsReview_(sheet, rowNumber, record, source, file, reason, raw) {
  let fileUrl = '';
  try {
    fileUrl = file.getUrl();
  } catch (err) {
    fileUrl = '';
  }

  appendNote_(sheet, rowNumber, 'Document classification needs review (' + source + '): ' + file.getName() + ' - ' + reason);
  logEvent_('document_classification_needs_review', record.company_name, record.month, 'Document classification needs review', {
    source: source,
    fileName: file.getName(),
    fileUrl: fileUrl,
    reason: reason,
    classification: raw
  });
}

function saveDriveFileForDocs_(file, folder, matchedDocs, moveSingleFile) {
  if (matchedDocs.length === 1) {
    renameToCanonical_(file, folder, matchedDocs[0], moveSingleFile);
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

function renameToCanonical_(file, folder, doc, moveFile) {
  const extension = getFileExtension_(file.getName()) || extensionFromMimeType_(file.getMimeType()) || splitExtensions_(doc.accepted_extensions)[0] || 'file';
  if (!extension) {
    return;
  }
  const canonicalName = doc.doc_key + '.' + extension;
  if (file.getName() === canonicalName) {
    if (moveFile) {
      moveFileToFolder_(file, folder);
    }
    return;
  }
  trashExistingFileNamed_(folder, canonicalName, file.getId());
  file.setName(canonicalName);
  if (moveFile) {
    moveFileToFolder_(file, folder);
  }
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

function uniqueValues_(values) {
  return values.filter(function (value, index, list) {
    return value && list.indexOf(value) === index;
  });
}
