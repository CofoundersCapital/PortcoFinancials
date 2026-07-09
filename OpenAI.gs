function openaiExtract(prompt) {
  const payload = {
    model: getConfigString_('OPENAI_MODEL'),
    instructions: 'Extract portfolio company financial metrics from messy source text. Use null when a value cannot be determined confidently. Do not guess.',
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text: prompt
          }
        ]
      }
    ],
    max_output_tokens: getConfigInteger_('OPENAI_MAX_OUTPUT_TOKENS'),
    store: false,
    text: {
      format: {
        type: 'json_schema',
        name: 'flash_report_extraction',
        description: 'Structured monthly portfolio company flash report extraction.',
        strict: true,
        schema: getFlashReportExtractionSchema_()
      }
    }
  };

  return fetchOpenAIJson_(payload, 'Missing OPENAI_API_KEY script property. Set it in Apps Script Project Settings before generating flash reports.');
}

function openaiClassifyDocument_(classificationInput) {
  const candidateDocs = classificationInput.candidateDocs || [];
  if (candidateDocs.length === 0) {
    return {
      matches: [],
      unmatched_reason: 'No configured required document types accept this file extension.'
    };
  }

  const payload = {
    model: getConfigString_('OPENAI_MODEL'),
    instructions: [
      'Classify portfolio company monthly reporting documents.',
      'Return every required document type that is clearly present in the file.',
      'A single file can satisfy multiple document types, such as financials, model, and forecast in different spreadsheet tabs.',
      'Assess every candidate doc_key exactly once. Use only the provided doc_key values. Use no matches when evidence is weak or ambiguous.'
    ].join(' '),
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_text',
            text: buildDocumentClassificationPrompt_(classificationInput)
          }
        ]
      }
    ],
    max_output_tokens: 1024,
    store: false,
    text: {
      format: {
        type: 'json_schema',
        name: 'document_classification',
        description: 'Required monthly reporting document type classification.',
        strict: true,
        schema: getDocumentClassificationSchema_(candidateDocs)
      }
    }
  };

  return fetchOpenAIJson_(payload, 'Missing OPENAI_API_KEY script property. Set it in Apps Script Project Settings before classifying uploads.');
}

function fetchOpenAIJson_(payload, missingApiKeyMessage) {
  const apiKey = getScriptProperty_(PROPERTY_KEYS.OPENAI_API_KEY);
  if (!apiKey) {
    throw new Error(missingApiKeyMessage);
  }

  const response = UrlFetchApp.fetch(getConfigString_('OPENAI_ENDPOINT'), {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('OpenAI API error ' + status + ': ' + text);
  }

  const parsed = JSON.parse(text);
  const answer = getOpenAIResponseText_(parsed);
  if (!answer) {
    throw new Error('OpenAI response did not include output text: ' + text);
  }

  return JSON.parse(answer);
}

function buildDocumentClassificationPrompt_(classificationInput) {
  const candidateDocs = classificationInput.candidateDocs || [];
  const missingDocs = classificationInput.missingDocs || [];
  const missingDocKeys = missingDocs.map(function (doc) {
    return doc.doc_key;
  });
  const docsText = candidateDocs.map(function (doc) {
    return [
      '- doc_key: ' + doc.doc_key,
      '  display_name: ' + doc.display_name,
      '  accepted_extensions: ' + (doc.accepted_extensions || 'any'),
      '  current_status: ' + (missingDocKeys.indexOf(doc.doc_key) >= 0 ? 'missing' : 'already_received'),
      '  guidance: ' + getDocumentClassificationGuidance_(doc)
    ].join('\n');
  }).join('\n');

  return [
    'Classify this received file against the configured required document checklist.',
    '',
    'Rules:',
    '- Return all matching doc_key values that are clearly present.',
    '- In assessments, include exactly one row for every candidate required doc.',
    '- Return multiple matches when one spreadsheet/workbook/PDF contains multiple required materials.',
    '- Set is_match to true only for high- or medium-confidence matches.',
    '- Return an empty matches array if the file is supporting material, ambiguous, or not one of the candidate required documents.',
    '- Do not infer a match only because a document is still missing.',
    '',
    'File metadata:',
    'filename: ' + (classificationInput.fileName || ''),
    'mime_type: ' + (classificationInput.mimeType || ''),
    'extension: ' + (classificationInput.extension || ''),
    '',
    'Candidate required docs:',
    docsText || '[none]',
    '',
    'Missing docs before this upload: ' + (missingDocKeys.join(', ') || '[none]'),
    '',
    'Extracted text preview:',
    classificationInput.extractedText || '[No text could be extracted.]'
  ].join('\n');
}

function getDocumentClassificationGuidance_(doc) {
  if (doc.doc_key === 'financials') {
    return 'Historical financial statements such as P&L, income statement, balance sheet, cash flow statement, monthly actuals, or financial statement package.';
  }
  if (doc.doc_key === 'model') {
    return 'Operating or financial model with assumptions, drivers, scenario logic, runway model, or multi-period business model tabs.';
  }
  if (doc.doc_key === 'forecast') {
    return 'Forecast, budget, plan, projections, budget-vs-actual update, or forward-looking financial update.';
  }
  return 'Use the display name and checklist context to decide whether this required material is present.';
}

function getOpenAIResponseText_(response) {
  if (response.output_text) {
    return String(response.output_text).trim();
  }

  const output = response.output || [];
  for (let i = 0; i < output.length; i++) {
    const item = output[i];
    const content = item.content || [];
    for (let j = 0; j < content.length; j++) {
      if (content[j].type === 'output_text' && content[j].text) {
        return String(content[j].text).trim();
      }
    }
  }

  return '';
}

function getDocumentClassificationSchema_(candidateDocs) {
  const docKeys = candidateDocs.map(function (doc) {
    return doc.doc_key;
  });

  return {
    type: 'object',
    additionalProperties: false,
    required: ['matches', 'assessments', 'unmatched_reason'],
    properties: {
      matches: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['doc_key', 'confidence', 'reason'],
          properties: {
            doc_key: {
              type: 'string',
              enum: docKeys
            },
            confidence: {
              type: 'string',
              enum: ['high', 'medium', 'low']
            },
            reason: { type: 'string' }
          }
        }
      },
      assessments: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['doc_key', 'confidence', 'is_match', 'reason'],
          properties: {
            doc_key: {
              type: 'string',
              enum: docKeys
            },
            confidence: {
              type: 'string',
              enum: ['high', 'medium', 'low', 'none']
            },
            is_match: { type: 'boolean' },
            reason: { type: 'string' }
          }
        }
      },
      unmatched_reason: { type: ['string', 'null'] }
    }
  };
}

function getFlashReportExtractionSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    required: [
      'revenue_month',
      'revenue_currency',
      'gross_margin_pct',
      'ebitda_month',
      'net_income_month',
      'cash_balance',
      'burn_rate_monthly',
      'runway_months',
      'budget_vs_actual_revenue_pct',
      'forecast_summary',
      'key_risks',
      'model_assumption_changes',
      'exceptions_for_partner_attention',
      'extraction_confidence',
      'extraction_notes'
    ],
    properties: {
      revenue_month: { type: ['number', 'null'] },
      revenue_currency: { type: 'string' },
      gross_margin_pct: { type: ['number', 'null'] },
      ebitda_month: { type: ['number', 'null'] },
      net_income_month: { type: ['number', 'null'] },
      cash_balance: { type: ['number', 'null'] },
      burn_rate_monthly: { type: ['number', 'null'] },
      runway_months: { type: ['number', 'null'] },
      budget_vs_actual_revenue_pct: { type: ['number', 'null'] },
      forecast_summary: { type: ['string', 'null'] },
      key_risks: {
        type: 'array',
        items: { type: 'string' }
      },
      model_assumption_changes: {
        type: 'array',
        items: { type: 'string' }
      },
      exceptions_for_partner_attention: {
        type: 'array',
        items: { type: 'string' }
      },
      extraction_confidence: {
        type: 'string',
        enum: ['high', 'medium', 'low']
      },
      extraction_notes: { type: 'string' }
    }
  };
}
