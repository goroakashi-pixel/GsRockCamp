function getSheet_() {
  var spreadsheet = getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + SHEET_NAME);
  }
  return sheet;
}

function buildHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) { map[normalizeHeader_(header)] = index; });
  return map;
}

function buildOriginalChordCell_(bar, part, barNumber) {
  var first = sanitizeChordInput_(bar.firstHalf);
  var second = sanitizeChordInput_(bar.secondHalf);
  validateBarPayload_(bar, part, barNumber);
  return first && second ? first + '│' + second : first;
}

function validateDraftMetadata_(metadata) {
  if (!metadata || typeof metadata !== 'object') {
    throw new Error('metadata が不正です。');
  }
  var originalKey = sanitizeKeyInput_(metadata.originalKey);
  var quizKey = sanitizeKeyInput_(metadata.quizKey);
  var youtubeUrl = sanitizeUrlInput_(metadata.youtubeUrl);

  if (!originalKey) {
    throw new Error('original_key を入力してください。');
  }
  if (!/^[A-G](?:#|b)?m?$/.test(originalKey)) {
    throw new Error('original_key の形式が不正です: ' + originalKey);
  }
  if (quizKey && !/^(C|Am)$/.test(quizKey)) {
    throw new Error('Quiz_key は C または Am で指定してください。');
  }
  if (youtubeUrl && !isSupportedYoutubeValue_(youtubeUrl)) {
    throw new Error('YouTube は動画URL / youtu.be / 11文字ID を入力してください。');
  }
}

function fetchSongBundle_(baseId) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) {
    throw new Error('DB シートにデータがありません。');
  }

  var headers = values[0];
  var headerMap = buildHeaderMap_(headers);
  var idColumn = findColumnIndex_(headers, headerMap, ['id']);
  if (idColumn < 0) {
    throw new Error('ID 列が見つかりません。');
  }

  var rowMap = {};
  for (var r = 1; r < values.length; r += 1) {
    var rowId = normalizeNumericString_(values[r][idColumn]);
    if (rowId) {
      rowMap[rowId] = { rowNumber: r + 1, values: values[r] };
    }
  }

  var rows = [];
  for (var offset = 0; offset < PART_ORDER.length; offset += 1) {
    var targetId = String(baseId + offset);
    if (!rowMap[targetId]) {
      throw new Error('指定IDの4行が揃っていません。欠損ID: ' + targetId);
    }
    rows.push(parseSheetRow_(rowMap[targetId], headers, headerMap, targetId, offset));
  }

  var partMap = {};
  rows.forEach(function(row, index) {
    var partKey = normalizePartKey_(row.part);
    if (partMap[partKey]) {
      partKey = PART_ORDER[index];
      row.part = partKey;
      row.partLabel = partKey;
    }
    partMap[partKey] = row;
  });

  PART_ORDER.forEach(function(partKey, index) {
    if (!partMap[partKey]) {
      rows[index].part = partKey;
      rows[index].partLabel = partKey;
      partMap[partKey] = rows[index];
    }
  });

  return { baseId: baseId, sheet: sheet, headers: headers, headerMap: headerMap, rows: rows, partMap: partMap };
}

function buildDraftMetadata_(bundle) {
  var first = bundle.rows[0];
  var originalKey = sanitizeKeyInput_(first.originalKey);
  var youtubeUrl = sanitizeUrlInput_(first.youtubeUrl);
  var notes = [];
  var status = {
    originalKeySource: 'sheet',
    youtubeSource: 'sheet',
    youtubeEmbeddable: !!buildYoutubeEmbedUrl_(youtubeUrl)
  };

  if (!originalKey) {
    originalKey = inferOriginalKeyFromBundle_(bundle);
    status.originalKeySource = originalKey ? 'inferred' : 'empty';
    notes.push(originalKey
      ? 'DB original_key は空欄です。既存コードから仮置きしました。読み込み後の候補を確認してください。'
      : 'DB original_key は空欄です。「AIでoriginal_Chord取得」で候補を作成してください。');
  }
  if (!youtubeUrl) {
    status.youtubeSource = 'manual_search';
    notes.push('DB YouTube は空欄です。「AIでoriginal_Chord取得」で候補を作成してください。');
  }

  return {
    originalKey: originalKey || '',
    quizKey: determineQuizKey_(originalKey),
    youtubeUrl: youtubeUrl,
    youtubeEmbedUrl: buildYoutubeEmbedUrl_(youtubeUrl),
    youtubeSearchUrl: buildYoutubeSearchUrl_(first.artist, first.title),
    status: status,
    notes: notes
  };
}

function normalizeHeader_(value) { return String(value || '').trim().toLowerCase().replace(/[\s_\-]/g, '').replace(/[^a-z0-9\u3040-\u30ff\u4e00-\u9fff]/g, ''); }

function validateBarPayload_(bar, part, barNumber) {
  if (!bar || typeof bar !== 'object') throw new Error(part + ' ' + barNumber + '小節 のデータが不正です。');
  var first = sanitizeChordInput_(bar.firstHalf);
  var second = sanitizeChordInput_(bar.secondHalf);
  if (!first && second) throw new Error(part + ' ' + barNumber + '小節 は後半のみ入力できません。');
  [first, second].forEach(function(chordText) {
    if (!chordText) return;
    if (chordText.indexOf('│') >= 0) throw new Error(part + ' ' + barNumber + '小節 は UI が自動で │ を付与します。');
    validateChordSymbol_(chordText, part, barNumber);
  });
}

function upsertCell_(headers, row, updates, skipped, updatedKeys, headerName, nextValue, protectExisting) {
  if (headerName === 'File') return;
  var column = headers.indexOf(headerName);
  if (column < 0) return;
  var currentValue = String(getCellByHeader_(row, headers, headerName) || '');
  var nextText = nextValue == null ? '' : String(nextValue);

  if (protectExisting && SAVE_MODE === 'EMPTY_ONLY' && currentValue && currentValue !== nextText) {
    skipped.push(headerName + ' は既存値ありのため未更新');
    return;
  }
  if (currentValue === nextText) return;
  updates[headerName] = nextText;
  updatedKeys.push(headerName);
}

function getBarColumn_(values, headers, headerMap, baseName, index) {
  var aliases = [baseName + index, baseName + '_' + index, baseName.toLowerCase() + '_' + index, baseName.toUpperCase() + '_' + index];
  var column = findColumnIndex_(headers, headerMap, aliases);
  return column >= 0 ? values[column] : '';
}

function buildSongResponse_(bundle, extra) {
  var first = bundle.rows[0];
  var draftMetadata = buildDraftMetadata_(bundle);
  var aiStatus = getOpenAiConfig_(true);
  var response = {
    ok: true,
    action: 'song',
    appVersion: APP_VERSION,
    baseId: bundle.baseId,
    idRange: [bundle.baseId, bundle.baseId + 1, bundle.baseId + 2, bundle.baseId + 3],
    connection: {
      connected: true,
      spreadsheetId: SPREADSHEET_ID,
      sheetName: SHEET_NAME,
      rowCount: bundle.sheet.getLastRow()
    },
    metadata: {
      data: toBooleanLikeString_(first.dataValue),
      artist: first.artist,
      title: first.title,
      rank: first.rank,
      era: first.era,
      youtubeUrl: first.youtubeUrl,
      youtubeEmbedUrl: buildYoutubeEmbedUrl_(first.youtubeUrl),
      originalKey: first.originalKey,
      quizKey: first.quizKey,
      draftOriginalKey: draftMetadata.originalKey,
      draftQuizKey: draftMetadata.quizKey,
      draftYoutubeUrl: draftMetadata.youtubeUrl,
      draftYoutubeEmbedUrl: draftMetadata.youtubeEmbedUrl,
      youtubeSearchUrl: draftMetadata.youtubeSearchUrl,
      ufretRawLines: [],
      ufretRawText: '',
      ufretStatus: 'idle',
      ufretSourceUrl: '',
      draftStatus: draftMetadata.status,
      draftNotes: draftMetadata.notes,
      aiStatus: aiStatus
    },
    parts: PART_ORDER.map(function(partKey) {
      return serializePartRow_(bundle.partMap[partKey]);
    }),
    logs: buildLoadLogs_(bundle, draftMetadata)
  };

  if (extra) {
    Object.keys(extra).forEach(function(key) {
      response[key] = extra[key];
    });
  }
  return response;
}

function buildLoadLogs_(bundle, draftMetadata) {
  var logs = [
    'App version: ' + APP_VERSION,
    'Spreadsheet 接続: OK (' + SPREADSHEET_ID + ')',
    'Sheet 確認: ' + SHEET_NAME,
    '取得ID: ' + bundle.baseId + '〜' + (bundle.baseId + 3),
    '取得行: ' + bundle.rows.map(function(row) { return row.rowNumber; }).join(', ')
  ];
  var aiStatus = getOpenAiConfig_(true);
  logs.push(aiStatus.configured
    ? 'AI original_Chord: READY (' + aiStatus.model + ')'
    : 'AI original_Chord: OPENAI_API_KEY 未設定（コード下書きはOFF。公開Web調査のみ）');
  if (draftMetadata.originalKey) logs.push('original_key 下書き: ' + draftMetadata.originalKey + ' [' + draftMetadata.status.originalKeySource + ']');
  if (draftMetadata.youtubeUrl) logs.push('YouTube 下書き: ' + draftMetadata.youtubeUrl + ' [' + draftMetadata.status.youtubeSource + ']');
  else logs.push('YouTube 下書き: なし（AI取得で公開Web調査を実行してください）。');
  return logs;
}

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('payload が不正です。');
  }
  normalizeBaseId_(payload.baseId);
  validateDraftMetadata_(payload.metadata);

  if (!Array.isArray(payload.parts) || payload.parts.length !== PART_ORDER.length) {
    throw new Error('parts は intro / A / B / サビ の4件が必要です。');
  }

  var seen = {};
  payload.parts.forEach(function(partPayload) {
    if (!partPayload || typeof partPayload !== 'object') {
      throw new Error('part payload が不正です。');
    }
    var canonicalPart = normalizePartKey_(partPayload.part);
    if (seen[canonicalPart]) {
      throw new Error('part が重複しています: ' + canonicalPart);
    }
    seen[canonicalPart] = true;

    if (!Array.isArray(partPayload.bars) || partPayload.bars.length !== BAR_COUNT) {
      throw new Error(canonicalPart + ' の bars は8件必要です。');
    }
    partPayload.bars.forEach(function(bar, index) {
      validateBarPayload_(bar, canonicalPart, index + 1);
    });
  });

  PART_ORDER.forEach(function(part) {
    if (!seen[part]) {
      throw new Error('part が不足しています: ' + part);
    }
  });
}

function buildErrorResponse_(error, context) {
  return {
    ok: false,
    action: context && context.action ? context.action : 'error',
    baseId: context && context.baseId ? String(context.baseId) : '',
    appVersion: APP_VERSION,
    connection: { connected: false, spreadsheetId: SPREADSHEET_ID, sheetName: SHEET_NAME },
    error: error && error.message ? error.message : String(error),
    logs: ['ERROR: ' + (error && error.message ? error.message : String(error))]
  };
}

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw new Error('Spreadsheet open failed: ' + error.message);
  }
}

function parseSheetRow_(rowInfo, headers, headerMap, idText, offset) {
  var values = rowInfo.values;
  var row = {
    id: idText,
    rowNumber: rowInfo.rowNumber,
    artist: getByAliases_(values, headers, headerMap, ['artist']),
    title: getByAliases_(values, headers, headerMap, ['title', 'song']),
    rank: getByAliases_(values, headers, headerMap, ['rank']),
    era: getByAliases_(values, headers, headerMap, ['era']),
    dataValue: getByAliases_(values, headers, headerMap, ['data']),
    part: getByAliases_(values, headers, headerMap, ['part']) || PART_ORDER[offset],
    partLabel: getByAliases_(values, headers, headerMap, ['part']) || PART_ORDER[offset],
    youtubeUrl: getByAliases_(values, headers, headerMap, ['youtube', 'youtubeurl']),
    originalKey: getByAliases_(values, headers, headerMap, ['originalkey']),
    quizKey: getByAliases_(values, headers, headerMap, ['quizkey', 'key']),
    originalChords: [],
    changeChords: [],
    degrees: [],
    functions: [],
    nonDiatonic: [],
    sheetValues: values
  };

  for (var i = 1; i <= BAR_COUNT; i += 1) {
    row.originalChords.push(getBarColumn_(values, headers, headerMap, 'originalchord', i));
    row.changeChords.push(getBarColumn_(values, headers, headerMap, 'changechord', i));
    row.degrees.push(getBarColumn_(values, headers, headerMap, 'degree', i));
    row.functions.push(getBarColumn_(values, headers, headerMap, 'function', i));
    row.nonDiatonic.push(getBarColumn_(values, headers, headerMap, 'nondiatonic', i));
  }
  return row;
}

function getByAliases_(values, headers, headerMap, aliases) {
  var column = findColumnIndex_(headers, headerMap, aliases);
  return column >= 0 ? values[column] : '';
}

function normalizeDraftMetadata_(metadata, bundle) {
  var inferredOriginalKey = inferOriginalKeyFromBundle_(bundle);
  var originalKey = sanitizeKeyInput_(metadata && metadata.originalKey) || inferredOriginalKey || 'C';
  var quizKey = sanitizeKeyInput_(metadata && metadata.quizKey) || determineQuizKey_(originalKey);
  var youtubeUrl = sanitizeUrlInput_(metadata && metadata.youtubeUrl) || sanitizeUrlInput_(bundle.rows[0].youtubeUrl);
  return {
    originalKey: originalKey,
    quizKey: quizKey,
    youtubeUrl: youtubeUrl,
    youtubeEmbedUrl: buildYoutubeEmbedUrl_(youtubeUrl),
    youtubeSearchUrl: buildYoutubeSearchUrl_(bundle.rows[0].artist, bundle.rows[0].title)
  };
}

function flushUpdates_(sheet, headers, rowUpdates) {
  rowUpdates.forEach(function(rowUpdate) {
    Object.keys(rowUpdate.updates).forEach(function(headerName) {
      var column = headers.indexOf(headerName);
      if (column < 0) throw new Error('更新対象の列が見つかりません: ' + headerName);
      sheet.getRange(rowUpdate.rowNumber, column + 1).setValue(rowUpdate.updates[headerName]);
    });
  });
}

function findColumnIndex_(headers, headerMap, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    var key = normalizeHeader_(aliases[i]);
    if (headerMap.hasOwnProperty(key)) return headerMap[key];
  }
  for (var j = 0; j < headers.length; j += 1) {
    var normalizedHeader = normalizeHeader_(headers[j]);
    for (var k = 0; k < aliases.length; k += 1) {
      if (normalizedHeader === normalizeHeader_(aliases[k])) return j;
    }
  }
  return -1;
}

function buildRowUpdate_(bundle, row, originalChords, theory, draftMeta) {
  var updates = {};
  var skipped = [];
  var updatedKeys = [];

  upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'original_key', draftMeta.originalKey, true);
  upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'Quiz_key', draftMeta.quizKey, true);
  upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'YouTube', draftMeta.youtubeUrl, true);
  upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'Part', normalizePartLabel_(row.part), true);

  for (var i = 1; i <= BAR_COUNT; i += 1) {
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'original_Chord_' + i, originalChords[i - 1], true);
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'change_Chord_' + i, theory.changeChord[i - 1], SAVE_MODE === 'EMPTY_ONLY');
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'degree_' + i, theory.degree[i - 1], SAVE_MODE === 'EMPTY_ONLY');
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'Function_' + i, theory.functionText[i - 1], SAVE_MODE === 'EMPTY_ONLY');
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'non_diatonic_' + i, theory.nonDiatonic[i - 1], SAVE_MODE === 'EMPTY_ONLY');
  }

  return { rowNumber: row.rowNumber, part: normalizePartKey_(row.part), updates: updates, skipped: skipped, updatedKeys: updatedKeys };
}

function serializePartRow_(row) {
  return {
    id: row.id,
    rowNumber: row.rowNumber,
    part: normalizePartKey_(row.part),
    partLabel: row.partLabel,
    originalKey: row.originalKey,
    quizKey: row.quizKey,
    bars: row.originalChords.map(function(cell, index) {
      var split = splitChordCell_(cell);
      return {
        bar: index + 1,
        combined: cell || '',
        firstHalf: split.first,
        secondHalf: split.second,
        aiSuggestion: split.aiSuggestion,
        aiDraftFirstHalf: '',
        aiDraftSecondHalf: '',
        changeChord: row.changeChords[index] || '',
        degree: row.degrees[index] || '',
        functionText: row.functions[index] || '',
        nonDiatonic: row.nonDiatonic[index] || ''
      };
    })
  };
}

function getCellByHeader_(row, headers, headerName) {
  var index = headers.indexOf(headerName);
  return index < 0 ? '' : (row.sheetValues && row.sheetValues[index] != null ? row.sheetValues[index] : '');
}

