function openaiExtract(prompt) {
  const apiKey = getScriptProperty_(PROPERTY_KEYS.OPENAI_API_KEY);
  if (!apiKey) {
    throw new Error('Missing OPENAI_API_KEY script property. Set it in Apps Script Project Settings before generating flash reports.');
  }

  const payload = {
    model: CONFIG.OPENAI_MODEL,
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
    max_output_tokens: CONFIG.OPENAI_MAX_OUTPUT_TOKENS,
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

  const response = UrlFetchApp.fetch(CONFIG.OPENAI_ENDPOINT, {
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
