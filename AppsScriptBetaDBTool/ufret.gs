function normalizeSuggestRequest_(input) {
  if (typeof input === 'object' && input) {
    return {
      baseId: input.baseId,
      referenceUrl: sanitizeUrlInput_(input.referenceUrl),
      ufretRawLines: Array.isArray(input.ufretRawLines) ? input.ufretRawLines : [],
      ufretRawText: String(input.ufretRawText || '')
    };
  }
  return {
    baseId: input,
    referenceUrl: '',
    ufretRawLines: [],
    ufretRawText: ''
  };
}

function normalizeUfretRawPayload_(lines, text) {
  var normalizedLines = Array.isArray(lines) ? lines.map(function(line) { return String(line || '').trim(); }).filter(Boolean) : [];
  var rawText = String(text || '').trim();
  if (!rawText && normalizedLines.length) rawText = normalizedLines.join('\n');
  if (!normalizedLines.length && rawText) normalizedLines = rawText.split(/\r?\n/).map(function(line) { return String(line || '').trim(); }).filter(Boolean);
  return {
    lines: normalizedLines,
    text: rawText,
    previewLines: normalizedLines.slice(0, 120)
  };
}

function normalizeUfretExtensionResponse_(payload, referenceUrl) {
  var sourceUrl = sanitizeUrlInput_((payload && payload.songUrl) || referenceUrl);
  var lines = Array.isArray(payload && payload.ufretRawLines)
    ? payload.ufretRawLines.map(function(line) { return String(line || '').trim(); }).filter(Boolean)
    : [];
  var previewText = String(payload && payload.previewText || '').trim();
  if (!previewText && lines.length) previewText = lines.join('\n');
  if (!lines.length && previewText) lines = previewText.split(/\r?\n/).map(function(line) { return String(line || '').trim(); }).filter(Boolean);
  return {
    ok: !!(payload && payload.ok),
    message: String((payload && payload.message) || ''),
    sourceUrl: sourceUrl,
    songId: String((payload && payload.songId) || ''),
    apiUrl: String((payload && payload.apiUrl) || ''),
    lineCount: Number((payload && payload.lineCount) || lines.length || 0),
    ufretRawLines: lines,
    ufretRawText: previewText,
    uniqueChords: Array.isArray(payload && payload.uniqueChords) ? payload.uniqueChords.slice(0, 400) : [],
    rawJsonText: String((payload && payload.rawJsonText) || ''),
    logText: String((payload && payload.logText) || ''),
    elapsedMs: Number((payload && payload.elapsedMs) || 0)
  };
}
