function claudeExtract(prompt) {
  const apiKey = getScriptProperty_(PROPERTY_KEYS.ANTHROPIC_API_KEY);
  if (!apiKey) {
    throw new Error('Missing ANTHROPIC_API_KEY script property. Set it in Apps Script Project Settings before generating flash reports.');
  }

  const payload = {
    model: CONFIG.ANTHROPIC_MODEL,
    max_tokens: 4096,
    messages: [
      {
        role: 'user',
        content: prompt
      }
    ]
  };

  const response = UrlFetchApp.fetch(CONFIG.ANTHROPIC_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': CONFIG.ANTHROPIC_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('Anthropic API error ' + status + ': ' + text);
  }

  const parsed = JSON.parse(text);
  const answer = (parsed.content || [])
    .filter(function (part) {
      return part.type === 'text';
    })
    .map(function (part) {
      return part.text;
    })
    .join('\n')
    .trim();

  return JSON.parse(extractJsonObject_(answer));
}

function extractJsonObject_(text) {
  const value = String(text || '').trim();
  if (value.charAt(0) === '{' && value.charAt(value.length - 1) === '}') {
    return value;
  }

  const first = value.indexOf('{');
  const last = value.lastIndexOf('}');
  if (first >= 0 && last > first) {
    return value.slice(first, last + 1);
  }

  throw new Error('Claude response did not contain a JSON object: ' + value);
}
