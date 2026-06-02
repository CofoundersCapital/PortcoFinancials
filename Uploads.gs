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

    fileIds.forEach(function (fileId) {
      processUploadedFile_(fileId, targetFolder, doc);
    });

    setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_status', 'received');
    setCellByHeader_(sheet, rowNumber, 'doc_' + doc.doc_key + '_date', new Date());
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
    let rowUpdated = false;

    getMissingDocs_(record, docs).forEach(function (doc) {
      const matchedFile = findMatchingDocumentFile_(folder, doc);
      if (!matchedFile) {
        return;
      }

      renameToCanonical_(matchedFile, doc);
      setCellByHeader_(sheet, entry.rowNumber, 'doc_' + doc.doc_key + '_status', 'received');
      setCellByHeader_(sheet, entry.rowNumber, 'doc_' + doc.doc_key + '_date', new Date());
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
  const file = DriveApp.getFileById(fileId);
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
}

function processSupportingFiles_(fileIds, targetFolder) {
  (fileIds || []).forEach(function (fileId, index) {
    const file = DriveApp.getFileById(fileId);
    const extension = getFileExtension_(file.getName());
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd-HHmmss');
    const name = 'supporting-' + timestamp + '-' + (index + 1) + (extension ? '.' + extension : '');
    trashExistingFileNamed_(targetFolder, name, file.getId());
    file.setName(name);
    moveFileToFolder_(file, targetFolder);
  });
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

function buildDocFilenamePattern_(doc) {
  const terms = [doc.doc_key].concat(String(doc.display_name || '').split(/[^a-zA-Z0-9]+/));
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
