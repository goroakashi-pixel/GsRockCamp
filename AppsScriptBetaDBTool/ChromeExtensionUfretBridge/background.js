chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (!message || message.action !== 'extractUfretPreviewFromReferenceUrl') return;
  extractUfretPreviewFromReferenceUrl(message.referenceUrl || '')
    .then((result) => sendResponse(result))
    .catch((error) => {
      sendResponse({ ok: false, message: error && error.message ? error.message : String(error), logText: '' });
    });
  return true;
});

async function extractUfretPreviewFromReferenceUrl(referenceUrl) {
  const logs = [];
  const startedAt = Date.now();
  const songUrl = String(referenceUrl || '').trim();
  logs.push(stamp('U-FRET Preview Extractor bridge(background) ver1.0.1'));
  logs.push(stamp('start'));
  logs.push(stamp('songUrl: ' + songUrl));

  const songId = extractSongId(songUrl);
  logs.push(stamp('songId: ' + songId));

  const apiUrl = 'https://www.ufret.jp/web_api/get_chord.php?song_id=' + encodeURIComponent(songId);
  logs.push(stamp('apiUrl: ' + apiUrl));

  const headers = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'X-Requested-With': 'XMLHttpRequest'
  };

  const response = await fetch(apiUrl, {
    method: 'GET',
    credentials: 'include',
    headers
  });

  logs.push(stamp('response status: ' + response.status));
  logs.push(stamp('response ok: ' + response.ok));
  logs.push(stamp('response content-type: ' + (response.headers.get('content-type') || '')));

  const rawText = await response.text();
  logs.push(stamp('response body length: ' + rawText.length));
  logs.push(stamp('response raw preview: ' + truncate(rawText, 500)));

  if (!response.ok) {
    return finish(false, { message: 'HTTP ' + response.status, songUrl, songId, apiUrl, previewText: '', rawJsonText: rawText, lineCount: 0, ufretRawLines: [], uniqueChords: [], logs, startedAt });
  }
  if (!rawText.trim()) {
    return finish(false, { message: 'API returned empty body', songUrl, songId, apiUrl, previewText: '', rawJsonText: rawText, lineCount: 0, ufretRawLines: [], uniqueChords: [], logs, startedAt });
  }

  logs.push(stamp('payload parse start'));
  const lines = parsePayload(rawText, logs);
  logs.push(stamp('payload parse success'));
  logs.push(stamp('lineCount: ' + lines.length));

  const previewText = buildCopyPastePreviewText(lines);
  const chordSequence = extractBracketChords(lines);
  const uniqueChords = dedupeKeepOrder(chordSequence);
  logs.push(stamp('chordSequence count: ' + chordSequence.length));
  logs.push(stamp('uniqueChords: ' + uniqueChords.join(', ')));

  return finish(true, {
    message: 'success',
    songUrl,
    songId,
    apiUrl,
    lineCount: lines.length,
    previewText,
    ufretRawLines: lines,
    rawJsonText: JSON.stringify(lines, null, 2),
    uniqueChords,
    logs,
    startedAt
  });
}

function finish(ok, payload) {
  return {
    ok,
    message: payload.message || '',
    songUrl: payload.songUrl || '',
    songId: payload.songId || '',
    apiUrl: payload.apiUrl || '',
    lineCount: payload.lineCount || 0,
    previewText: payload.previewText || '',
    ufretRawLines: payload.ufretRawLines || [],
    rawJsonText: payload.rawJsonText || '',
    uniqueChords: payload.uniqueChords || [],
    elapsedMs: Date.now() - payload.startedAt,
    logText: buildLogText(payload.logs || [])
  };
}

function extractSongId(url) {
  const match = String(url || '').match(/[?&]data=(\d+)/i);
  if (!match) throw new Error('song_id(data=...) を URL から取得できません。');
  return match[1];
}

function parsePayload(rawText, logs) {
  const text = String(rawText || '').trim();
  if (!text) throw new Error('empty payload');
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      logs.push(stamp('parse mode: JSON array'));
      return parsed.map(normalizeLine);
    }
    logs.push(stamp('parse mode: JSON parsed but not array'));
  } catch (error) {
    logs.push(stamp('parse mode JSON array failed: ' + error.message));
  }

  const quoted = [];
  const regex = /"(?:\\.|[^"\\])*"/g;
  let match;
  while ((match = regex.exec(text))) {
    try { quoted.push(normalizeLine(JSON.parse(match[0]))); } catch (_ignore) {}
  }
  if (quoted.length) {
    logs.push(stamp('parse mode: quoted strings'));
    return quoted;
  }

  const lines = text.split(/\r?\n/).map(normalizeLine).filter(Boolean);
  if (lines.length) {
    logs.push(stamp('parse mode: line split'));
    return lines;
  }

  throw new Error('payload parse failed');
}

function buildCopyPastePreviewText(lines) {
  return (lines || []).map((line, index) => `${index}: "${escapeForDisplay(line)}"`).join('\n');
}

function extractBracketChords(lines) {
  const result = [];
  const regex = /\[([^\[\]]+)\]/g;
  (lines || []).forEach((line) => {
    const text = String(line || '');
    let match;
    while ((match = regex.exec(text))) {
      const chord = String(match[1] || '').replace(/\u3000/g, ' ').replace(/\s+/g, '').trim();
      if (chord) result.push(chord);
    }
  });
  return result;
}

function normalizeLine(value) {
  return String(value == null ? '' : value)
    .replace(/^\uFEFF/, '')
    .replace(/\r/g, '')
    .trimEnd();
}

function escapeForDisplay(text) {
  return String(text == null ? '' : text)
    .replace(/\r/g, '\\r')
    .replace(/\n/g, '\\n')
    .replace(/"/g, '\\"');
}

function dedupeKeepOrder(list) {
  const seen = new Set();
  const result = [];
  (list || []).forEach((item) => {
    const key = String(item || '');
    if (!key || seen.has(key)) return;
    seen.add(key);
    result.push(key);
  });
  return result;
}

function truncate(text, maxLength) {
  const value = String(text == null ? '' : text);
  return value.length <= maxLength ? value : value.slice(0, maxLength) + ' ...';
}

function stamp(message) {
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  return `[${hh}:${mm}:${ss}] ${message}`;
}

function buildLogText(logs) {
  return (logs || []).join('\n');
}
