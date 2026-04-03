const SPREADSHEET_ID = '1D6d0iNhMdZn8I0Jj-m1tAT0blZIgec4qTHlAJuGo4ms';
const SHEET_NAME = 'DB';
const PART_ORDER = ['intro', 'A', 'B', 'サビ'];
const SAVE_MODE = 'EMPTY_ONLY'; // 'EMPTY_ONLY' | 'FORCE'
const BAR_COUNT = 8;
const APP_VERSION = '1.1.4';
const OPENAI_MODEL = 'gpt-5-mini';
const ENABLE_RULE_BASED_FALLBACK = false;
const AI_JSON_RETRY_MAX = 2;
const REQUIRED_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/script.external_request'
];

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.appVersion = APP_VERSION;
  return template.evaluate()
    .setTitle('GoRockCamp DB投入支援ツール β v' + APP_VERSION)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw new Error('Spreadsheet open failed: ' + error.message);
  }
}

function getSheet_() {
  var spreadsheet = getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + SHEET_NAME);
  }
  return sheet;
}

function loadSong(baseId) {
  try {
    var normalizedBaseId = normalizeBaseId_(baseId);
    var bundle = fetchSongBundle_(normalizedBaseId);
    var response = buildSongResponse_(bundle, {
      ok: true,
      action: 'loadSong',
      message: '読み込み成功: ID ' + normalizedBaseId + '〜' + (normalizedBaseId + 3),
      stage: 'draft_ready'
    });
    return response;
  } catch (error) {
    return buildErrorResponse_(error, { action: 'loadSong', baseId: baseId });
  }
}

function checkAuthorizationStatus() {
  try {
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    var status = String(authInfo.getAuthorizationStatus());
    var authorizationUrl = String(authInfo.getAuthorizationUrl() || '');
    return {
      ok: true,
      action: 'checkAuthorizationStatus',
      appVersion: APP_VERSION,
      requiredScopes: REQUIRED_SCOPES.slice(),
      authorized: status === String(ScriptApp.AuthorizationStatus.NOT_REQUIRED),
      authorizationStatus: status,
      authorizationUrl: authorizationUrl,
      message: status === String(ScriptApp.AuthorizationStatus.NOT_REQUIRED)
        ? '権限OK'
        : '未承認の権限があります。承認URLを開いて承認してください。'
    };
  } catch (error) {
    return buildErrorResponse_(error, { action: 'checkAuthorizationStatus' });
  }
}

function suggestOriginalChords(baseId) {
  try {
    var normalizedBaseId = normalizeBaseId_(baseId);
    var bundle = fetchSongBundle_(normalizedBaseId);
    var response = buildSongResponse_(bundle, {
      ok: true,
      action: 'suggestOriginalChords',
      stage: 'ai_chord_drafted'
    });
    var aiConfig = getOpenAiConfig_(true);
    var research = hydrateDraftMetadataFromWeb_(bundle, response, { metadataOnly: false });
    response.logs = (response.logs || []).concat([
      'AI precheck: OPENAI_API_KEY ' + (aiConfig.configured ? 'configured' : 'missing')
    ]);
    if (!aiConfig.configured) {
      response.stage = 'metadata_only';
      if (ENABLE_RULE_BASED_FALLBACK) {
        var research = hydrateDraftMetadataFromWeb_(bundle, response, { metadataOnly: false });
        var heuristicDraft = buildRuleBasedDraftFromResearch_(bundle, research);
        if (heuristicDraft) {
          applyAiDraftToResponse_(response, heuristicDraft);
          response.message = 'OPENAI_API_KEY 未設定のため、公開Web調査ベースで original_Chord の仮下書きを生成しました。';
          response.logs = (response.logs || []).concat(['AI fallback draft: web chord scrape から生成しました。']);
        }
      }
      if (!response.message) response.message = buildResearchOnlyMessage_(research);
      response.logs = (response.logs || []).concat(buildResearchOnlyLogs_(research));
      return response;
    }

    try {
      var draft = requestAiOriginalChordDraft_(bundle, aiConfig);
      var aiCellCount = countNonEmptyDraftCells_(draft.partMap);
      if (aiCellCount <= 0) {
        throw new Error('AI returned empty bars');
      }
      applyAiDraftToResponse_(response, draft);
      var appliedCount = countAppliedAiDraftCellsInResponse_(response.parts);
      if (appliedCount <= 0) {
        throw new Error('UI apply failure: response.parts へAI下書きを反映できませんでした');
      }
      response.message = 'AI original_Chord 下書きを取得し、' + appliedCount + 'セルをUI表示用に反映しました。';
      response.logs = (response.logs || []).concat(draft.logs || []);
      response.logs.push('AI original_Chord 下書き成功: ' + appliedCount + 'セルを response.parts に反映');
    } catch (aiError) {
      response.stage = 'metadata_only';
      if (aiError && Array.isArray(aiError.draftLogs)) {
        response.logs = (response.logs || []).concat(aiError.draftLogs);
      }
      var fallbackResearch = hydrateDraftMetadataFromWeb_(bundle, response, { metadataOnly: false });
      var fallbackDraft = buildRuleBasedDraftFromResearch_(bundle, fallbackResearch);
      var fallbackCount = countNonEmptyDraftCells_(fallbackDraft && fallbackDraft.partMap);
      if (fallbackDraft && fallbackCount > 0) {
        applyAiDraftToResponse_(response, fallbackDraft);
        var appliedFallbackCount = countAppliedAiDraftCellsInResponse_(response.parts);
        if (appliedFallbackCount > 0) {
          response.stage = 'ai_chord_drafted';
          response.message = 'AI返答は空でしたが、公開情報の補完下書きを' + appliedFallbackCount + 'セル分表示しました。';
          response.logs = (response.logs || []).concat([
            'AI returned empty bars',
            'fallback applied: research.sources から叩き台を生成',
            'AI original_Chord 下書き成功(補完): ' + appliedFallbackCount + 'セルを response.parts に反映'
          ]);
          return response;
        }
      }
      var failureType = classifyAiDraftError_(aiError);
      response.message = 'AI返答は受信しましたが、コード下書きは空のため失敗しました。';
      response.logs = (response.logs || []).concat([
        failureType + ': ' + aiError.message,
        'AI返答は受信したが、コード下書きは空のため失敗',
        'UI apply skipped: AIドラフトは反映していません'
      ]);
    }
    return response;
  } catch (error) {
    return buildErrorResponse_(error, { action: 'suggestOriginalChords', baseId: baseId });
  }
}

function saveSong(payload) {
  try {
    validatePayload_(payload);

    var normalizedBaseId = normalizeBaseId_(payload.baseId);
    var bundle = fetchSongBundle_(normalizedBaseId);
    var draftMeta = normalizeDraftMetadata_(payload.metadata, bundle);
    var updates = [];
    var skipped = [];

    payload.parts.forEach(function(partPayload) {
      var canonicalPart = normalizePartKey_(partPayload.part);
      var targetRow = bundle.partMap[canonicalPart];
      if (!targetRow) {
        throw new Error('保存対象のパートが見つかりません: ' + partPayload.part);
      }

      var rebuiltBars = [];
      for (var i = 1; i <= BAR_COUNT; i += 1) {
        rebuiltBars.push(buildOriginalChordCell_(partPayload.bars[i - 1], canonicalPart, i));
      }

      var theory = recalcTheory_(rebuiltBars, draftMeta.originalKey, draftMeta.quizKey);
      var rowUpdate = buildRowUpdate_(bundle, targetRow, rebuiltBars, theory, draftMeta);
      updates.push(rowUpdate);
      skipped = skipped.concat(rowUpdate.skipped);
    });

    flushUpdates_(bundle.sheet, bundle.headers, updates);

    var reloadedBundle = fetchSongBundle_(normalizedBaseId);
    var response = buildSongResponse_(reloadedBundle, {
      ok: true,
      action: 'saveSong',
      message: '確定保存成功: original_key / YouTube / original_Chord / 理論列を再読込しました',
      saveMode: SAVE_MODE,
      skipped: skipped,
      stage: 'saved'
    });

    response.saveSummary = updates.map(function(item) {
      return {
        part: item.part,
        rowNumber: item.rowNumber,
        updatedCells: item.updatedKeys,
        skipped: item.skipped
      };
    });

    return response;
  } catch (error) {
    return buildErrorResponse_(error, { action: 'saveSong', baseId: payload && payload.baseId });
  }
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

function flushUpdates_(sheet, headers, rowUpdates) {
  rowUpdates.forEach(function(rowUpdate) {
    Object.keys(rowUpdate.updates).forEach(function(headerName) {
      var column = headers.indexOf(headerName);
      if (column < 0) throw new Error('更新対象の列が見つかりません: ' + headerName);
      sheet.getRange(rowUpdate.rowNumber, column + 1).setValue(rowUpdate.updates[headerName]);
    });
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

function getCellByHeader_(row, headers, headerName) {
  var index = headers.indexOf(headerName);
  return index < 0 ? '' : (row.sheetValues && row.sheetValues[index] != null ? row.sheetValues[index] : '');
}

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

function buildOriginalChordCell_(bar, part, barNumber) {
  var first = sanitizeChordInput_(bar.firstHalf);
  var second = sanitizeChordInput_(bar.secondHalf);
  validateBarPayload_(bar, part, barNumber);
  return first && second ? first + '│' + second : first;
}

function validateChordSymbol_(chordText, part, barNumber) {
  var slashCount = (String(chordText).match(/\//g) || []).length;
  if (slashCount > 1) throw new Error(part + ' ' + barNumber + '小節 に不正な / 連続表記があります: ' + chordText);
  var base = String(chordText).split('/')[0];
  if (!/^([A-G](?:#|b|♭)?)(?:[0-9A-Za-z+()#♭b\-susdimaugmajMadd]*)$/.test(base)) throw new Error(part + ' ' + barNumber + '小節 のコード形式が不正です: ' + chordText);
  if (slashCount === 1) {
    var bass = String(chordText).split('/')[1] || '';
    if (!/^[A-G](?:#|b|♭)?$/.test(bass)) throw new Error(part + ' ' + barNumber + '小節 のオンコード表記が不正です: ' + chordText);
  }
}

function splitChordCell_(cellText) {
  var text = String(cellText || '').trim();
  if (!text) return { first: '', second: '', aiSuggestion: [] };
  var parts = text.split('│').map(function(item) { return item.trim(); }).filter(Boolean);
  return { first: parts[0] || '', second: parts[1] || '', aiSuggestion: parts };
}

function recalcTheory_(originalChords, originalKey, quizKey) {
  var sourceKey = normalizeKeyName_(originalKey || quizKey || 'C');
  var targetKey = normalizeKeyName_(quizKey || originalKey || 'C');
  var shift = getSemitoneShift_(sourceKey, targetKey);
  return {
    changeChord: originalChords.map(function(cell) { return transposeCell_(cell, shift); }),
    degree: originalChords.map(function(cell) { return buildTheoryCell_(cell, function(chord) { return getDegreeText_(chord, targetKey); }); }),
    functionText: originalChords.map(function(cell) { return buildTheoryCell_(cell, function(chord) { return getFunctionText_(chord, targetKey); }); }),
    nonDiatonic: originalChords.map(function(cell) { return buildNonDiatonicCell_(cell, targetKey); })
  };
}

function buildTheoryCell_(cellText, mapper) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  return parts.length ? parts.map(mapper).join('│') : '';
}

function buildNonDiatonicCell_(cellText, keyName) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  if (!parts.length) return '';
  return parts.some(function(chord) { return !isChordDiatonic_(chord, keyName); }) ? 'TRUE' : 'FALSE';
}

function transposeCell_(cellText, shift) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  return parts.length ? parts.map(function(chord) { return transposeChordSymbol_(chord, shift); }).join('│') : '';
}

function transposeChordSymbol_(chordText, shift) {
  var parsed = parseChordSymbol_(chordText);
  var root = transposePitchName_(parsed.root, shift);
  var bass = parsed.bass ? transposePitchName_(parsed.bass, shift) : '';
  return root + parsed.quality + (bass ? '/' + bass : '');
}

function getDegreeText_(chordText, keyName) {
  var parsed = parseChordSymbol_(chordText);
  return getRomanDegree_(parsed.root, keyName) + inferDegreeSuffix_(parsed.quality);
}

function getFunctionText_(chordText, keyName) {
  var parsed = parseChordSymbol_(chordText);
  var roman = getRomanDegree_(parsed.root, keyName);
  if (/Ⅴ|Ⅶ/.test(roman)) return isChordDiatonic_(chordText, keyName) ? 'D' : 'セカンダリードミナント';
  if (/Ⅱ|Ⅳ/.test(roman)) return isChordDiatonic_(chordText, keyName) ? 'SD' : '同主調借用';
  if (/Ⅰ|Ⅵ|Ⅲ/.test(roman)) return isChordDiatonic_(chordText, keyName) ? 'T' : 'モーダルインターチェンジ';
  return 'クロマチック';
}

function isChordDiatonic_(chordText, keyName) {
  var scaleSet = getScaleSetForKey_(keyName);
  return buildChordPitchClasses_(chordText).every(function(note) { return scaleSet[note]; });
}

function buildChordPitchClasses_(chordText) {
  var parsed = parseChordSymbol_(chordText);
  var rootIndex = noteNameToIndex_(parsed.root);
  if (rootIndex < 0) return [];
  var intervals = [0, 4, 7];
  if (/^m(?!aj)/.test(parsed.quality)) intervals = [0, 3, 7];
  if (/dim/.test(parsed.quality)) intervals = [0, 3, 6];
  if (/aug/.test(parsed.quality)) intervals = [0, 4, 8];
  if (/sus4/.test(parsed.quality)) intervals = [0, 5, 7];
  if (/6/.test(parsed.quality)) intervals.push(9);
  if (/maj7|M7|mM7/.test(parsed.quality)) intervals.push(11); else if (/7/.test(parsed.quality)) intervals.push(10);
  if (/9|add9/.test(parsed.quality)) intervals.push(14);
  if (parsed.bass) intervals.push(noteNameToIndex_(parsed.bass) - rootIndex);

  var pitchSet = {};
  intervals.forEach(function(interval) { pitchSet[indexToSharpName_((rootIndex + interval) % 12)] = true; });
  return Object.keys(pitchSet);
}

function getScaleSetForKey_(keyName) {
  var normalized = normalizeKeyName_(keyName);
  if (/m$/.test(normalized) && normalized !== 'Am') return getNaturalMinorScaleSet_(normalized.replace(/m$/, ''));
  if (normalized === 'Am') return getNaturalMinorScaleSet_('A');
  return getMajorScaleSet_(normalized || 'C');
}

function getMajorScaleSet_(keyName) {
  var root = noteNameToIndex_(keyName);
  var set = {};
  [0, 2, 4, 5, 7, 9, 11].forEach(function(interval) { set[indexToSharpName_((root + interval) % 12)] = true; });
  return set;
}

function getNaturalMinorScaleSet_(keyName) {
  var root = noteNameToIndex_(keyName);
  var set = {};
  [0, 2, 3, 5, 7, 8, 10].forEach(function(interval) { set[indexToSharpName_((root + interval) % 12)] = true; });
  return set;
}

function getRomanDegree_(rootName, keyName) {
  var tonic = normalizeKeyRoot_(keyName || 'C');
  var diff = ((noteNameToIndex_(rootName) - noteNameToIndex_(tonic)) % 12 + 12) % 12;
  var map = {0:'Ⅰ',1:'bⅡ',2:'Ⅱ',3:'bⅢ',4:'Ⅲ',5:'Ⅳ',6:'bⅤ',7:'Ⅴ',8:'bⅥ',9:'Ⅵ',10:'bⅦ',11:'Ⅶ'};
  return map[diff] || '?';
}

function inferDegreeSuffix_(quality) {
  quality = String(quality || '');
  if (!quality) return '';
  if (/m7-5/.test(quality)) return 'm7-5';
  if (/maj7|M7/.test(quality)) return 'M7';
  if (/m6/.test(quality)) return 'm6';
  if (/6/.test(quality)) return '6';
  if (/m7/.test(quality)) return 'm7';
  if (/m(?!aj)/.test(quality)) return 'm';
  if (/dim/.test(quality)) return 'dim';
  if (/aug/.test(quality)) return 'aug';
  if (/sus4/.test(quality)) return 'sus4';
  if (/add9/.test(quality)) return '(add9)';
  if (/9/.test(quality)) return '9';
  if (/7/.test(quality)) return '7';
  return quality;
}

function parseChordSymbol_(chordText) {
  var halves = String(chordText || '').trim().split('/');
  var main = halves[0] || '';
  var bass = halves[1] || '';
  var match = main.match(/^([A-G](?:#|b|♭)?)(.*)$/);
  if (!match) throw new Error('コードを解釈できません: ' + chordText);
  return { root: normalizeKeyName_(match[1]), quality: match[2] || '', bass: bass ? normalizeKeyName_(bass) : '' };
}

function inferOriginalKeyFromBundle_(bundle) {
  var chords = [];
  bundle.rows.forEach(function(row) {
    row.originalChords.forEach(function(cell) {
      splitChordCell_(cell).aiSuggestion.forEach(function(chord) { chords.push(chord); });
    });
  });
  return inferOriginalKeyFromChordList_(chords);
}

function inferOriginalKeyFromChordList_(chords) {
  if (!chords || !chords.length) return '';
  var candidates = ['C','G','D','A','E','F','Bb','Eb','Am','Em','Dm','Bm','Cm','Fm'];
  var bestKey = '';
  var bestScore = -1;
  candidates.forEach(function(candidate) {
    var score = 0;
    chords.forEach(function(chord, index) {
      var parsed = parseChordSymbol_(chord);
      var roman = getRomanDegree_(parsed.root, candidate);
      if (isChordDiatonic_(chord, candidate)) score += 3;
      if (/Ⅰ|Ⅳ|Ⅴ|Ⅵ/.test(roman)) score += 2;
      if (index === 0 && /^(Ⅰ|Ⅵ)/.test(roman)) score += 1;
      if (index === chords.length - 1 && /^(Ⅰ|Ⅵ)/.test(roman)) score += 2;
    });
    if (score > bestScore) { bestScore = score; bestKey = candidate; }
  });
  return bestKey;
}

function determineQuizKey_(originalKey) {
  var normalized = sanitizeKeyInput_(originalKey);
  return normalized && /m$/.test(normalized) ? 'Am' : 'C';
}

function getSemitoneShift_(fromKey, toKey) {
  return ((noteNameToIndex_(normalizeKeyRoot_(toKey)) - noteNameToIndex_(normalizeKeyRoot_(fromKey))) % 12 + 12) % 12;
}

function normalizeKeyRoot_(keyName) {
  var normalized = normalizeKeyName_(keyName || 'C');
  if (normalized === 'Am') return 'A';
  return normalized.replace(/m$/, '');
}

function transposePitchName_(noteName, shift) {
  var index = noteNameToIndex_(noteName);
  return index < 0 ? noteName : indexToSharpName_((index + shift) % 12);
}

function noteNameToIndex_(noteName) {
  var normalized = normalizeKeyName_(noteName);
  var map = {C:0,'B#':0,'C#':1,Db:1,D:2,'D#':3,Eb:3,E:4,Fb:4,'E#':5,F:5,'F#':6,Gb:6,G:7,'G#':8,Ab:8,A:9,'A#':10,Bb:10,B:11,Cb:11};
  return map.hasOwnProperty(normalized) ? map[normalized] : -1;
}

function indexToSharpName_(index) { return ['C','C#','D','Eb','E','F','F#','G','Ab','A','Bb','B'][((index % 12) + 12) % 12]; }
function normalizeKeyName_(text) { return String(text || '').trim().replace(/♭/g, 'b').replace(/＃/g, '#').replace(/([A-Ga-g])/, function(match){ return match.toUpperCase(); }); }
function sanitizeKeyInput_(value) { return normalizeKeyName_(value).replace(/major$/i, '').replace(/minor$/i, 'm'); }
function sanitizeUrlInput_(value) { return String(value == null ? '' : value).trim(); }

function normalizePartKey_(partText) {
  var normalized = String(partText || '').trim();
  var map = { intro:'intro', 'イントロ':'intro', A:'A', 'Aメロ':'A', B:'B', 'Bメロ':'B', 'サビ':'サビ', chorus:'サビ' };
  return map.hasOwnProperty(normalized) ? map[normalized] : normalized;
}

function normalizePartLabel_(partText) {
  var canonical = normalizePartKey_(partText);
  var map = { intro:'intro', A:'Aメロ', B:'Bメロ', 'サビ':'サビ' };
  return map[canonical] || canonical;
}

function normalizeBaseId_(baseId) {
  var text = normalizeNumericString_(baseId);
  if (!text) throw new Error('先頭IDは数値で入力してください。');
  return Number(text);
}

function normalizeNumericString_(value) { var text = String(value == null ? '' : value).trim(); return /^\d+$/.test(text) ? text : ''; }
function sanitizeChordInput_(value) { return String(value == null ? '' : value).replace(/\s+/g, ' ').trim().replace(/｜/g, '│'); }
function toBooleanLikeString_(value) { var text = String(value == null ? '' : value).trim().toUpperCase(); return text === 'TRUE' || text === 'FALSE' ? text : (text || ''); }
function isSupportedYoutubeValue_(rawValue) { return !!buildYoutubeEmbedUrl_(rawValue); }

function buildYoutubeEmbedUrl_(rawValue) {
  var text = String(rawValue || '').trim();
  if (!text) return '';
  var idMatch = text.match(/(?:v=|youtu\.be\/|embed\/|shorts\/)([A-Za-z0-9_-]{11})/);
  var videoId = idMatch ? idMatch[1] : text;
  return /^[A-Za-z0-9_-]{11}$/.test(videoId) ? 'https://www.youtube.com/embed/' + videoId : '';
}

function buildYoutubeSearchUrl_(artist, title) {
  var query = [artist || '', title || '', 'official'].join(' ').trim();
  return query ? 'https://www.youtube.com/results?search_query=' + encodeURIComponent(query) : '';
}

function getByAliases_(values, headers, headerMap, aliases) {
  var column = findColumnIndex_(headers, headerMap, aliases);
  return column >= 0 ? values[column] : '';
}

function getBarColumn_(values, headers, headerMap, baseName, index) {
  var aliases = [baseName + index, baseName + '_' + index, baseName.toLowerCase() + '_' + index, baseName.toUpperCase() + '_' + index];
  var column = findColumnIndex_(headers, headerMap, aliases);
  return column >= 0 ? values[column] : '';
}

function buildHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) { map[normalizeHeader_(header)] = index; });
  return map;
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

function normalizeHeader_(value) { return String(value || '').trim().toLowerCase().replace(/[\s_\-]/g, '').replace(/[^a-z0-9\u3040-\u30ff\u4e00-\u9fff]/g, ''); }

function requestAiOriginalChordDraft_(bundle, config) {
  config = config && config.configured ? config : getOpenAiConfig_();
  var logs = [
    'OpenAI web_search 実行開始',
    'AI draft model: ' + config.model + ' (' + (config.modelSource || 'default') + ')'
  ];
  var lastError = null;
  var rawTextForRetry = '';
  var retryHint = '';
  for (var attempt = 1; attempt <= AI_JSON_RETRY_MAX; attempt += 1) {
    try {
      logs.push('AI JSON生成 attempt ' + attempt + '/' + AI_JSON_RETRY_MAX);
      var aiJson = runAiDraftAttempt_(bundle, config, attempt, rawTextForRetry, retryHint);
      rawTextForRetry = extractResponseText_(aiJson);
      var parsed = extractStructuredDraftObject_(aiJson, logs);
      validateAiDraftSchema_(parsed);
      var draft = normalizeAiDraftResponse_(parsed, bundle);
      var keyCheck = checkDraftBarsKeyAgainstOriginal_(draft, bundle);
      if (keyCheck.retry) {
        logs.push('AI bars key check: originalKey=' + keyCheck.originalKey + ' / suspectedBarsKey=' + keyCheck.suspectedBarsKey + ' / retry');
        throw new Error('AI bars appear transposed to Quiz_key');
      }
      logs.push('AI bars accepted in original key: ' + keyCheck.originalKey);
      var draftCount = countNonEmptyDraftCells_(draft.partMap);
      if (draftCount <= 0) {
        throw new Error('AI returned empty bars');
      }
      logs.push('AI下書き受領');
      logs.push('JSON検証OK');
      logs.push('non-empty draft count: ' + draftCount);
      draft.logs = logs;
      return draft;
    } catch (error) {
      lastError = error;
      var errText = String(error && error.message || '');
      if (/transposed to Quiz_key/i.test(errText)) {
        retryHint = '前回は C/Am 側へ転調されていました。今回は必ず original_key のまま返し、C/Amへ移調しないでください。';
      } else if (/empty bars/i.test(errText)) {
        retryHint = '前回は bars が空でした。各partに最低1小節以上、可能なら8小節埋めてください。';
      } else {
        retryHint = '前回はJSON整形に失敗しました。JSON以外を一切含めないでください。';
      }
      var parsePos = extractJsonErrorPosition_(error.message);
      var errorType = classifyAiDraftError_(error);
      logs.push(errorType + ' attempt ' + attempt + ': ' + error.message);
      logs.push('AI raw snippet: ' + summarizeRawText_(rawTextForRetry, 800));
      if (parsePos >= 0) logs.push('parse失敗位置: ' + parsePos);
      if (attempt >= AI_JSON_RETRY_MAX) break;
      logs.push('JSON再試行を実行します。');
    }
  }
  var wrapped = new Error('AI JSON整形失敗: ' + (lastError ? lastError.message : 'unknown'));
  wrapped.draftLogs = logs;
  throw wrapped;
}

function getOpenAiConfig_(allowMissing) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = String(props.getProperty('OPENAI_API_KEY') || '').trim();
  var scriptModel = String(props.getProperty('OPENAI_MODEL') || '').trim();
  var model = scriptModel || OPENAI_MODEL;
  var modelSource = scriptModel ? 'script_properties' : 'default_mini';
  if (!apiKey) {
    if (allowMissing) return { configured: false, apiKey: '', model: model || OPENAI_MODEL, modelSource: modelSource };
    throw new Error('OPENAI_API_KEY が Script Properties に未設定です。AI original_Chord 取得は設定後に利用してください。');
  }
  return { configured: true, apiKey: apiKey, model: model || OPENAI_MODEL, modelSource: modelSource };
}

function callOpenAiResponsesApiRaw_(apiKey, payload) {
  var response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload)
  });
  var status = response.getResponseCode();
  var body = response.getContentText();
  if (status < 200 || status >= 300) throw new Error('OpenAI API error (' + status + '): ' + body);
  return JSON.parse(body);
}

function extractResponseText_(json) {
  if (json.output_text) return json.output_text;
  var chunks = [];
  (json.output || []).forEach(function(item) {
    (item.content || []).forEach(function(content) {
      if (content.text) chunks.push(content.text);
      if (content.type === 'output_text' && content.text) chunks.push(content.text);
    });
  });
  if (!chunks.length) throw new Error('OpenAI response から本文を取得できませんでした。');
  return chunks.join('\n');
}

function buildChordResearchPrompt_(bundle, previousRawText) {
  var retryHint = arguments.length > 2 ? String(arguments[2] || '').trim() : '';
  var first = bundle.rows[0];
  var retrySection = previousRawText ? [
    '前回出力がJSONスキーマ不一致でした。JSON以外を一切含めず、schema準拠オブジェクトのみ返してください。',
    '--- previous output start ---',
    previousRawText.slice(0, 4000),
    '--- previous output end ---',
    retryHint
  ].join('\n') : '';
  if (!retrySection && retryHint) retrySection = retryHint;
  return [
    '公開情報を web_search で調査し、original_key / YouTube / intro,A,B,サビ の original_chord 下書きを返してください。',
    '参照優先順位: 1) 公式YouTube 2) U-FRET/ChordWiki 3) その他公開Web情報',
    'Artist: ' + first.artist,
    'Title: ' + first.title,
    'Current DB original_key: ' + (first.originalKey || ''),
    'Current DB YouTube: ' + (first.youtubeUrl || ''),
    '出力制約:',
    '- part は intro / A / B / サビ',
    '- bars は各 part 8件',
    '- bars は必ず original_key（原曲キー）のコードで返すこと',
    '- Quiz_key(C/Am) への変換は禁止。C/Am へ移調したコードを返さないこと',
    '- return only original_Chord bars; do not output change_Chord/degree/function',
    '- Return chord bars only in the song original key. Never transpose bars to C or Am.',
    '- Quiz_key is only for later theory conversion, not for returned bars.',
    '- secondHalf only 禁止',
    '- slash はオンコード専用',
    '- 1小節2コード時だけ firstHalf/secondHalf を使用',
    '- 不明な箇所は空文字、推測で埋めすぎない',
    '- 公開情報が十分なら part 単位で整理して返す',
    '- 各 part に最低1小節以上は firstHalf を埋めること（全part空欄は不可）',
    '- 可能なら8小節すべて埋める。4小節反復が明確なら5〜8小節へ反復してよい',
    '- 前回が空barsだった場合は、曖昧でも公開情報から叩き台を構成して埋めること',
    '- notes は短文のみ（最大3件、URL/markdown禁止）',
    retrySection
  ].join('\n');
}

function buildAiResponseSchema_() {
  return {
    name: 'go_rock_camp_chord_draft',
    schema: {
      type: 'object',
      additionalProperties: false,
      required: ['originalKey', 'youtubeUrl', 'notes', 'parts'],
      properties: {
        originalKey: { type: 'string' },
        youtubeUrl: { type: 'string' },
        notes: { type: 'array', maxItems: 3, items: { type: 'string', maxLength: 120 } },
        parts: {
          type: 'array',
          items: {
            type: 'object',
            additionalProperties: false,
            required: ['part', 'bars'],
            properties: {
              part: { type: 'string', enum: PART_ORDER },
              bars: {
                type: 'array',
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['bar', 'firstHalf', 'secondHalf'],
                  properties: {
                    bar: { type: 'integer', minimum: 1, maximum: BAR_COUNT },
                    firstHalf: { type: 'string' },
                    secondHalf: { type: 'string' }
                  }
                }
              }
            }
          }
        }
      }
    }
  };
}

function runAiDraftAttempt_(bundle, config, attempt, previousRawText, retryHint) {
  return callOpenAiResponsesApiRaw_(config.apiKey, {
    model: config.model,
    tools: [{ type: 'web_search' }],
    input: buildChordResearchPrompt_(bundle, attempt > 1 ? previousRawText : '', attempt > 1 ? retryHint : ''),
    text: {
      format: {
        type: 'json_schema',
        name: buildAiResponseSchema_().name,
        schema: buildAiResponseSchema_().schema,
        strict: true
      }
    }
  });
}

function extractStructuredDraftObject_(json, logs) {
  if (json && json.output_parsed) {
    if (logs) logs.push('JSON受信: output_parsed を使用');
    return json.output_parsed;
  }
  var output = json && Array.isArray(json.output) ? json.output : [];
  for (var i = 0; i < output.length; i += 1) {
    var contentList = Array.isArray(output[i].content) ? output[i].content : [];
    for (var j = 0; j < contentList.length; j += 1) {
      if (contentList[j] && typeof contentList[j].parsed === 'object' && contentList[j].parsed) {
        if (logs) logs.push('JSON受信: content.parsed を使用');
        return contentList[j].parsed;
      }
    }
  }
  var text = extractResponseText_(json);
  try {
    if (logs) logs.push('JSON受信: output_text 全文を直接JSON.parse');
    return JSON.parse(text);
  } catch (error) {
    var extracted = extractFirstJsonObject_(text);
    if (logs) logs.push('JSON切り出し成功: length=' + extracted.__jsonLength + ' chars');
    return extracted;
  }
}

function validateAiDraftSchema_(raw) {
  if (!raw || typeof raw !== 'object') throw new Error('schema validation error: root object');
  if (!Array.isArray(raw.parts)) throw new Error('schema validation error: parts array');
  raw.parts.forEach(function(part) {
    if (!part || PART_ORDER.indexOf(part.part) < 0) throw new Error('schema validation error: invalid part');
    if (!Array.isArray(part.bars) || part.bars.length !== BAR_COUNT) throw new Error('schema validation error: bars length');
    part.bars.forEach(function(bar, index) {
      if (!bar || Number(bar.bar) !== index + 1) throw new Error('schema validation error: bar index');
      if (typeof bar.firstHalf !== 'string' || typeof bar.secondHalf !== 'string') throw new Error('schema validation error: bar cell type');
      if (!bar.firstHalf && bar.secondHalf) throw new Error('schema validation error: secondHalf only');
    });
  });
}

function classifyAiDraftError_(error) {
  var message = String((error && error.message) || error || '');
  if (/ui apply failure/i.test(message)) return 'UI apply failure';
  if (/transposed to Quiz_key/i.test(message)) return 'AI bars key mismatch';
  if (/empty bars/i.test(message)) return 'AI returned empty bars';
  if (/schema validation error/i.test(message)) return 'schema validation error';
  if (/json parse error|unexpected|json/i.test(message)) return 'JSON parse error';
  if (/openai api error/i.test(message)) return 'OpenAI API error';
  return 'AI draft error';
}

function extractJsonErrorPosition_(message) {
  var match = String(message || '').match(/position\s+(\d+)/i);
  return match ? Number(match[1]) : -1;
}

function summarizeRawText_(text, maxLength) {
  var normalized = String(text || '').replace(/\s+/g, ' ').trim();
  if (!normalized) return '(empty)';
  return normalized.slice(0, maxLength || 300);
}

function hydrateDraftMetadataFromWeb_(bundle, response, options) {
  options = options || {};
  var metadata = response.metadata || {};
  var needsResearch = !sanitizeKeyInput_(metadata.draftOriginalKey) || !sanitizeUrlInput_(metadata.draftYoutubeUrl) || !options.metadataOnly;
  if (!needsResearch) return { ok: true, skipped: true, reason: 'metadata_already_present', logs: ['公開Web調査: 既存metadataがあるためスキップ'] };

  try {
    var research = researchPublicSongData_(bundle, options);
    applyResearchMetadataToResponse_(response, research);
    response.logs = (response.logs || []).concat(research.logs || []);
    return research;
  } catch (error) {
    var analysis = analyzeResearchError_(error);
    response.metadata = response.metadata || {};
    response.metadata.draftNotes = filterResearchNotes_(response.metadata.draftNotes || []);
    response.metadata.draftNotes.push(analysis.note);
    response.logs = (response.logs || []).concat(analysis.logs);
    return {
      ok: false,
      skipped: true,
      permissionRequired: analysis.permissionRequired,
      reason: analysis.reason,
      note: analysis.note,
      logs: analysis.logs
    };
  }
}

function applyResearchMetadataToResponse_(response, research) {
  if (!research) return;
  response.metadata = response.metadata || {};
  response.metadata.draftStatus = response.metadata.draftStatus || {};
  response.metadata.draftNotes = filterResearchNotes_(response.metadata.draftNotes || []);

  if (research.originalKey && !sanitizeKeyInput_(response.metadata.draftOriginalKey)) {
    response.metadata.draftOriginalKey = research.originalKey;
    response.metadata.draftQuizKey = determineQuizKey_(research.originalKey);
    response.metadata.draftStatus.originalKeySource = research.originalKeySource || 'web_research';
  }

  if (research.youtubeUrl && !sanitizeUrlInput_(response.metadata.draftYoutubeUrl)) {
    response.metadata.draftYoutubeUrl = research.youtubeUrl;
    response.metadata.draftYoutubeEmbedUrl = buildYoutubeEmbedUrl_(research.youtubeUrl);
    response.metadata.draftStatus.youtubeSource = research.youtubeSource || 'web_research';
  }

  if (research.notes && research.notes.length) {
    response.metadata.draftNotes = response.metadata.draftNotes.concat(research.notes);
  }
}

function researchPublicSongData_(bundle, options) {
  options = options || {};
  var first = bundle.rows[0];
  var query = [first.artist || '', first.title || ''].join(' ').trim();
  if (!query) throw new Error('artist/title が空のため Web 調査できません。');

  var logs = ['Web research start: ' + query];
  var notes = [];
  var sources = [];
  var youtubeSearch = searchYoutubeCandidate_(first.artist, first.title);
  var youtubeResult = youtubeSearch.candidate;
  logs = logs.concat(youtubeSearch.logs || []);
  if (youtubeResult) {
    sources.push(youtubeResult);
    logs.push('YouTube candidate found: ' + youtubeResult.url);
  } else {
    logs.push('YouTube candidate not found by public search');
  }

  var chordSearch = findChordReferenceCandidates_(query);
  var chordResults = chordSearch.results || [];
  logs = logs.concat(chordSearch.logs || []);
  chordResults.forEach(function(result) { sources.push(result); });
  if (chordResults.length) logs.push('Chord reference candidates: ' + chordResults.length);
  else logs.push('Chord reference candidates: 0');

  var scrapedChordPool = [];
  chordResults.slice(0, 3).forEach(function(result) {
    try {
      var html = fetchHtmlText_(result.url);
      var chords = extractChordCandidatesFromHtml_(html);
      var keyHints = extractOriginalKeyHintsFromHtml_(html);
      result.chords = chords.slice(0, 64);
      result.keyHints = keyHints.slice(0, 8);
      if (result.chords.length) {
        scrapedChordPool = scrapedChordPool.concat(result.chords);
        logs.push('Chord scrape OK: ' + result.url + ' / chords=' + result.chords.length);
      } else {
        logs.push('Chord scrape empty: ' + result.url);
      }
      if (result.keyHints.length) logs.push('Key hints: ' + result.keyHints.join(', '));
    } catch (error) {
      logs.push('Chord scrape failed: ' + result.url + ' / ' + error.message);
    }
  });

  var originalKey = pickOriginalKeyFromSources_(sources, scrapedChordPool);
  var youtubeUrl = youtubeResult ? youtubeResult.url : '';
  if (originalKey) notes.push('公開Web調査から original_key 候補を取得しました。保存前に確認してください。');
  if (youtubeUrl) notes.push('公開Web調査から YouTube 候補を取得しました。公式動画か確認してください。');
  if (!originalKey) notes.push('公開Web調査では original_key 候補を確定できませんでした。必要なら手入力してください。');
  if (!youtubeUrl) notes.push('公開Web調査では YouTube 候補を取得できませんでした。検索リンクで確認してください。');
  if (!options.metadataOnly && !getOpenAiConfig_(true).configured) {
    logs.push('OPENAI_API_KEY 未設定のため original_Chord AI draft は未実行');
  }

  return {
    ok: true,
    originalKey: originalKey,
    originalKeySource: originalKey ? 'web_research' : '',
    youtubeUrl: youtubeUrl,
    youtubeSource: youtubeUrl ? 'web_research' : '',
    sources: sources,
    logs: logs,
    notes: notes
  };
}

function searchYoutubeCandidate_(artist, title) {
  var query = [artist || '', title || '', 'official YouTube'].join(' ').trim();
  if (!query) return { candidate: null, logs: [] };
  var logs = [];
  var searchUrls = buildSearchUrls_(query);
  for (var u = 0; u < searchUrls.length; u += 1) {
    try {
      var html = fetchHtmlText_(searchUrls[u].url);
      var results = extractSearchResultsFromHtml_(html);
      for (var i = 0; i < results.length; i += 1) {
        var url = decodeSearchRedirectUrl_(results[i].url);
        if (!/youtube\.com|youtu\.be/.test(url)) continue;
        var videoId = extractYoutubeVideoId_(url);
        if (videoId) {
          logs.push('YouTube search provider OK: ' + searchUrls[u].name);
          return {
            candidate: {
              type: 'youtube',
              title: results[i].title,
              snippet: results[i].snippet,
              url: 'https://www.youtube.com/watch?v=' + videoId
            },
            logs: logs
          };
        }
      }
      logs.push('YouTube search provider no match: ' + searchUrls[u].name);
    } catch (error) {
      logs.push('YouTube search provider failed: ' + searchUrls[u].name + ' / ' + error.message);
    }
  }
  return { candidate: null, logs: logs };
}

function findChordReferenceCandidates_(query) {
  var results = [];
  var logs = [];
  var searches = [
    query + ' site:chordwiki.jpn.org',
    query + ' site:ufret.jp',
    query + ' コード'
  ];

  searches.forEach(function(searchQuery) {
    var providers = buildSearchUrls_(searchQuery);
    providers.forEach(function(provider) {
      try {
        var html = fetchHtmlText_(provider.url);
        var beforeCount = results.length;
        extractSearchResultsFromHtml_(html).forEach(function(item) {
          var url = decodeSearchRedirectUrl_(item.url);
          pushChordReferenceCandidate_(results, url, item.title, item.snippet);
        });
        extractTargetLinksFromHtml_(html, ['ufret.jp', 'chordwiki.jpn.org']).forEach(function(url) {
          pushChordReferenceCandidate_(results, url, 'direct-link', provider.name + ' raw-link');
        });
        logs.push('Chord search provider OK: ' + provider.name + ' / ' + searchQuery + ' / +' + (results.length - beforeCount));
      } catch (error) {
        logs.push('Chord search provider failed: ' + provider.name + ' / ' + searchQuery + ' / ' + error.message);
      }
    });
  });

  return { results: results.filter(function(item) { return item.url; }).slice(0, 6), logs: logs };
}

function pushChordReferenceCandidate_(resultList, rawUrl, title, snippet) {
  var url = normalizeCandidateUrl_(decodeSearchRedirectUrl_(rawUrl));
  if (!/^https?:\/\//.test(url)) return;
  if (!/chordwiki\.jpn\.org|ufret\.jp/.test(url)) return;
  if (resultList.some(function(existing) { return existing.url === url; })) return;
  resultList.push({
    type: /ufret\.jp/.test(url) ? 'ufret' : 'chordwiki',
    title: title || '',
    snippet: snippet || '',
    url: url,
    chords: [],
    keyHints: []
  });
}

function fetchHtmlText_(url) {
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (compatible; GoRockCampBot/1.0; +https://script.google.com/)'
    }
  });
  var status = response.getResponseCode();
  if (status < 200 || status >= 300) throw new Error('HTTP ' + status);
  return response.getContentText();
}

function extractSearchResultsFromHtml_(html) {
  var text = String(html || '');
  var results = [];
  var regex = /<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?(?:<a[^>]+class="[^"]*result__snippet[^"]*"[^>]*>|<div[^>]+class="[^"]*result__snippet[^"]*"[^>]*>)([\s\S]*?)(?:<\/a>|<\/div>)/g;
  var match;
  while ((match = regex.exec(text)) && results.length < 12) {
    results.push({
      url: decodeHtmlEntities_(stripTags_(match[1])),
      title: decodeHtmlEntities_(stripTags_(match[2])),
      snippet: decodeHtmlEntities_(stripTags_(match[3]))
    });
  }
  if (!results.length) {
    var fallbackRegex = /<a[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g;
    while ((match = fallbackRegex.exec(text)) && results.length < 12) {
      var url = decodeHtmlEntities_(stripTags_(match[1]));
      if (url.indexOf('uddg=') < 0) continue;
      results.push({ url: url, title: decodeHtmlEntities_(stripTags_(match[2])), snippet: '' });
    }
  }
  return results;
}

function decodeDuckDuckGoRedirect_(url) {
  var text = String(url || '');
  var match = text.match(/[?&]uddg=([^&]+)/);
  if (match) return decodeURIComponent(match[1]);
  return text;
}

function decodeSearchRedirectUrl_(url) {
  var text = String(url || '');
  if (/duckduckgo\.com/.test(text) && /[?&]uddg=/.test(text)) return decodeDuckDuckGoRedirect_(text);
  if (/bing\.com\/ck\//.test(text)) {
    var targetMatch = text.match(/[?&]u=([^&]+)/);
    if (targetMatch) {
      try {
        var decoded = decodeURIComponent(targetMatch[1]);
        var cleaned = decoded.replace(/^a1/, '');
        var urlMatch = cleaned.match(/https?:\/\/.*/);
        if (urlMatch) return urlMatch[0];
      } catch (_error) {}
    }
  }
  return text;
}

function buildSearchUrls_(query) {
  var encoded = encodeURIComponent(query);
  return [
    { name: 'duckduckgo', url: 'https://duckduckgo.com/html/?q=' + encoded },
    { name: 'bing', url: 'https://www.bing.com/search?q=' + encoded }
  ];
}

function extractTargetLinksFromHtml_(html, domains) {
  var text = decodeHtmlEntities_(String(html || ''));
  var escaped = text.replace(/\\\//g, '/');
  var links = [];
  var regex = /(https?:\/\/[^\s"'<>\\)]+)/g;
  var match;
  while ((match = regex.exec(escaped)) && links.length < 120) {
    var url = match[1];
    if (!domains.some(function(domain) { return url.indexOf(domain) >= 0; })) continue;
    links.push(url);
  }
  return links;
}

function normalizeCandidateUrl_(url) {
  return String(url || '')
    .replace(/[),.;]+$/g, '')
    .replace(/&amp;/g, '&');
}

function stripTags_(html) {
  return String(html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function decodeHtmlEntities_(text) {
  return String(text || '')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

function extractChordCandidatesFromHtml_(html) {
  var text = decodeHtmlEntities_(String(html || '').replace(/\n/g, ' '));
  var matches = text.match(/\b[A-G](?:#|b)?(?:maj7|M7|m7-5|mM7|m7|m6|m|7|6|9|add9|sus4|dim|aug)?(?:\/[A-G](?:#|b)?)?\b/g) || [];
  var blacklist = { HTML:true, HTTP:true, HTTPS:true };
  var filtered = [];
  matches.forEach(function(chord) {
    if (blacklist[chord]) return;
    if (filtered.length && filtered[filtered.length - 1] === chord) return;
    filtered.push(chord);
  });
  return filtered;
}

function extractOriginalKeyHintsFromHtml_(html) {
  var text = decodeHtmlEntities_(stripTags_(html));
  var hints = [];
  var regex = /(?:原曲キー|Key|キー)\s*[:：]?\s*([A-G](?:#|b)?m?)/gi;
  var match;
  while ((match = regex.exec(text)) && hints.length < 8) {
    hints.push(sanitizeKeyInput_(match[1]));
  }
  return hints.filter(Boolean);
}

function pickOriginalKeyFromSources_(sources, chords) {
  var keyScores = {};
  sources.forEach(function(source) {
    (source.keyHints || []).forEach(function(keyName) {
      if (!keyName) return;
      keyScores[keyName] = (keyScores[keyName] || 0) + 3;
    });
  });
  var inferred = inferOriginalKeyFromChordList_(chords || []);
  if (inferred) keyScores[inferred] = (keyScores[inferred] || 0) + 2;
  var bestKey = '';
  var bestScore = -1;
  Object.keys(keyScores).forEach(function(keyName) {
    if (keyScores[keyName] > bestScore) {
      bestKey = keyName;
      bestScore = keyScores[keyName];
    }
  });
  return bestKey;
}

function extractYoutubeVideoId_(url) {
  var text = String(url || '').trim();
  var match = text.match(/(?:v=|youtu\.be\/|embed\/|shorts\/)([A-Za-z0-9_-]{11})/);
  return match ? match[1] : '';
}

function filterResearchNotes_(notes) {
  return (notes || []).filter(function(note) {
    return note.indexOf('読み込み時に公開Web調査を試みます。') < 0 &&
      note.indexOf('DB original_key は空欄です。既存コードから仮置きしました。') < 0 &&
      note.indexOf('DB YouTube は空欄です。') < 0 &&
      note.indexOf('公開Web調査に失敗しました。') < 0 &&
      note.indexOf('公開Web調査では original_key 候補を確定できませんでした。') < 0 &&
      note.indexOf('公開Web調査では YouTube 候補を取得できませんでした。') < 0;
  });
}

function analyzeResearchError_(error) {
  var message = error && error.message ? error.message : String(error);
  if (/script\.external_request|UrlFetchApp\.fetch.*権限|外部アクセス|external_request/i.test(message)) {
    return {
      permissionRequired: true,
      reason: 'external_request_auth_required',
      note: '外部アクセス権限が未承認のため公開Web調査を実行できませんでした。再デプロイ後に「外部サービスへの接続」を承認してください。',
      logs: [
        '公開Web調査は未実行（script.external_request 未承認）。',
        '外部アクセス権限が未承認です。権限確認ボタンから承認URLを開いてください。'
      ]
    };
  }
  return {
    permissionRequired: false,
    reason: 'research_failed',
    note: '公開Web調査に失敗しました。検索リンクまたは手入力で確認してください。',
    logs: ['公開Web調査エラー: ' + message]
  };
}

function buildResearchOnlyMessage_(research) {
  if (research && research.permissionRequired) {
    return '外部アクセス権限が未承認のため公開Web調査を実行できませんでした。再承認後は OPENAI_API_KEY 未設定でも公開Web調査は動作します。';
  }
  if (research && research.ok) {
    return '公開Web調査で original_key / YouTube 候補を更新しました。original_Chord の AI 下書きは Script Properties の OPENAI_API_KEY 設定後に利用できます。';
  }
  return '公開Web調査は候補取得まで完了しませんでした。検索リンクまたは手入力で補完してください。';
}

function buildResearchOnlyLogs_(research) {
  var logs = ['AI original_Chord 下書き: OPENAI_API_KEY 未設定のため未実行'];
  if (research && research.permissionRequired) {
    return logs;
  }
  if (research && research.ok) {
    logs.push('公開Web調査: 実行済み');
    return logs;
  }
  logs.push('公開Web調査: 候補取得なし');
  return logs;
}

function buildRuleBasedDraftFromResearch_(bundle, research) {
  if (!research || !research.ok) return null;
  var chordPool = extractChordPoolFromResearch_(research);
  if (!chordPool.length) return null;

  var draft = {
    originalKey: sanitizeKeyInput_(research.originalKey) || inferOriginalKeyFromChordList_(chordPool) || '',
    quizKey: '',
    youtubeUrl: sanitizeUrlInput_(research.youtubeUrl),
    notes: ['＜音寧コメ＞公開Web調査の抽出コードから仮下書きを作成しました。必ず人間確認で修正してください。'],
    partMap: {}
  };
  draft.quizKey = determineQuizKey_(draft.originalKey);

  PART_ORDER.forEach(function(part, partIndex) {
    draft.partMap[part] = [];
    var partSeed = chordPool.slice(partIndex * 4, partIndex * 4 + 8);
    if (!partSeed.length) partSeed = chordPool.slice(0, Math.min(8, chordPool.length));
    if (partSeed.length >= 4 && partSeed.length < 8) {
      while (partSeed.length < 8) {
        partSeed.push(partSeed[partSeed.length % 4]);
      }
    }
    for (var i = 0; i < BAR_COUNT; i += 1) {
      var poolIndex = i % Math.max(partSeed.length, 1);
      var chordText = partSeed[poolIndex] || chordPool[(partIndex * BAR_COUNT + i) % chordPool.length];
      draft.partMap[part].push({
        bar: i + 1,
        firstHalf: chordText || '',
        secondHalf: ''
      });
    }
  });

  return draft;
}

function extractChordPoolFromResearch_(research) {
  var primary = [];
  var secondary = [];
  (research.sources || []).forEach(function(source) {
    var chords = Array.isArray(source.chords) ? source.chords : [];
    if (!chords.length) return;
    var bucket = /\/song\.php/.test(source.url || '') ? primary : secondary;
    chords.forEach(function(chord) {
      var normalized = sanitizeChordInput_(chord);
      if (!normalized) return;
      try {
        validateChordSymbol_(normalized, 'research', 1);
        if (!bucket.some(function(existing) { return existing === normalized; })) bucket.push(normalized);
      } catch (_error) {}
    });
  });
  return primary.concat(secondary).slice(0, 64);
}

function extractFirstJsonObject_(text) {
  var raw = String(text || '');
  var start = raw.indexOf('{');
  if (start < 0) throw new Error('JSON parse error: JSON開始 "{" が見つかりません');

  var depth = 0;
  var inString = false;
  var escapeNext = false;
  var end = -1;

  for (var i = start; i < raw.length; i += 1) {
    var ch = raw.charAt(i);
    if (escapeNext) {
      escapeNext = false;
      continue;
    }
    if (inString) {
      if (ch === '\\') {
        escapeNext = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') {
      depth += 1;
      continue;
    }
    if (ch === '}') {
      depth -= 1;
      if (depth === 0) {
        end = i;
        break;
      }
      if (depth < 0) throw new Error('JSON parse error: 波かっこの対応が不正です');
    }
  }

  if (end < 0) throw new Error('JSON parse error: JSON終端 "}" を特定できません');
  var jsonText = raw.slice(start, end + 1);
  try {
    var parsed = JSON.parse(jsonText);
    if (parsed && typeof parsed === 'object') parsed.__jsonLength = jsonText.length;
    return parsed;
  } catch (error) {
    throw new Error('JSON parse error: ' + error.message);
  }
}

function normalizeAiDraftResponse_(raw, bundle) {
  var result = { originalKey: sanitizeKeyInput_(raw.originalKey), youtubeUrl: sanitizeUrlInput_(raw.youtubeUrl), notes: Array.isArray(raw.notes) ? raw.notes.map(sanitizeAiNote_).filter(Boolean) : [], partMap: {} };
  PART_ORDER.forEach(function(part) {
    result.partMap[part] = Array.from({ length: BAR_COUNT }, function(_, index) { return { bar: index + 1, firstHalf: '', secondHalf: '' }; });
  });
  (raw.parts || []).forEach(function(partInfo) {
    var canonical = normalizePartKey_(partInfo.part);
    if (!result.partMap[canonical]) return;
    (partInfo.bars || []).slice(0, BAR_COUNT).forEach(function(bar, idx) {
      result.partMap[canonical][idx] = { bar: idx + 1, firstHalf: sanitizeAiChordDraftValue_(bar.firstHalf), secondHalf: sanitizeAiChordDraftValue_(bar.secondHalf) };
      validateBarPayload_(result.partMap[canonical][idx], canonical, idx + 1);
    });
  });
  if (!result.originalKey) result.originalKey = inferOriginalKeyFromBundle_(bundle);
  result.quizKey = determineQuizKey_(result.originalKey);
  if (result.youtubeUrl && !isSupportedYoutubeValue_(result.youtubeUrl)) result.youtubeUrl = '';
  return result;
}

function sanitizeAiNote_(note) {
  var text = String(note == null ? '' : note).trim();
  text = text.replace(/https?:\/\/\S+/g, '').replace(/\[[^\]]+\]\([^)]+\)/g, '').trim();
  return text.slice(0, 120);
}

function sanitizeAiChordDraftValue_(value) { return sanitizeChordInput_(String(value == null ? '' : value).replace(/\|/g, '│')); }

function applyAiDraftToResponse_(response, draft) {
  if (draft.originalKey) {
    response.metadata.draftOriginalKey = draft.originalKey;
    response.metadata.draftQuizKey = draft.quizKey;
  }
  if (draft.youtubeUrl) {
    response.metadata.draftYoutubeUrl = draft.youtubeUrl;
    response.metadata.draftYoutubeEmbedUrl = buildYoutubeEmbedUrl_(draft.youtubeUrl);
  }
  if (draft.notes && draft.notes.length) {
    response.metadata.draftNotes = (response.metadata.draftNotes || []).concat(draft.notes);
  }
  response.parts.forEach(function(part) {
    var aiBars = draft.partMap[part.part] || [];
    part.bars.forEach(function(bar, index) {
      var aiBar = aiBars[index] || { firstHalf: '', secondHalf: '' };
      bar.aiDraftFirstHalf = aiBar.firstHalf || '';
      bar.aiDraftSecondHalf = aiBar.secondHalf || '';
    });
  });
}

function countNonEmptyDraftCells_(partMap) {
  if (!partMap || typeof partMap !== 'object') return 0;
  var count = 0;
  PART_ORDER.forEach(function(part) {
    var bars = partMap[part] || [];
    bars.forEach(function(bar) {
      if (!bar) return;
      if (sanitizeChordInput_(bar.firstHalf)) count += 1;
      if (sanitizeChordInput_(bar.secondHalf)) count += 1;
    });
  });
  return count;
}

function countAppliedAiDraftCellsInResponse_(parts) {
  if (!Array.isArray(parts)) return 0;
  var count = 0;
  parts.forEach(function(part) {
    (part.bars || []).forEach(function(bar) {
      if (sanitizeChordInput_(bar.aiDraftFirstHalf)) count += 1;
      if (sanitizeChordInput_(bar.aiDraftSecondHalf)) count += 1;
    });
  });
  return count;
}

function checkDraftBarsKeyAgainstOriginal_(draft, bundle) {
  var originalKey = sanitizeKeyInput_(draft.originalKey) || sanitizeKeyInput_(bundle.rows[0].originalKey) || inferOriginalKeyFromBundle_(bundle) || 'C';
  var quizKey = determineQuizKey_(originalKey);
  var chords = collectDraftChords_(draft.partMap);
  if (!chords.length) {
    return { retry: false, originalKey: originalKey, suspectedBarsKey: '' };
  }
  if (originalKey === 'C' || originalKey === 'Am') {
    return { retry: false, originalKey: originalKey, suspectedBarsKey: inferOriginalKeyFromChordList_(chords) || '' };
  }
  var suspectedBarsKey = inferOriginalKeyFromChordList_(chords) || '';
  var scoreOriginal = scoreDraftFitForKey_(chords, originalKey);
  var scoreQuiz = scoreDraftFitForKey_(chords, quizKey);
  var retry = suspectedBarsKey === quizKey && scoreQuiz >= scoreOriginal + 3;
  return {
    retry: retry,
    originalKey: originalKey,
    suspectedBarsKey: suspectedBarsKey || '(unknown)',
    scoreOriginal: scoreOriginal,
    scoreQuiz: scoreQuiz
  };
}

function collectDraftChords_(partMap) {
  var chords = [];
  PART_ORDER.forEach(function(part) {
    var bars = (partMap && partMap[part]) || [];
    bars.forEach(function(bar) {
      var first = sanitizeChordInput_(bar && bar.firstHalf);
      var second = sanitizeChordInput_(bar && bar.secondHalf);
      if (first) chords.push(first);
      if (second) chords.push(second);
    });
  });
  return chords;
}

function scoreDraftFitForKey_(chords, keyName) {
  var score = 0;
  (chords || []).forEach(function(chord, index) {
    try {
      var roman = getRomanDegree_(parseChordSymbol_(chord).root, keyName);
      if (isChordDiatonic_(chord, keyName)) score += 3;
      if (/Ⅰ|Ⅳ|Ⅴ|Ⅵ/.test(roman)) score += 2;
      if (index === 0 && /^(Ⅰ|Ⅵ)/.test(roman)) score += 1;
      if (index === (chords.length - 1) && /^(Ⅰ|Ⅵ)/.test(roman)) score += 2;
    } catch (_error) {}
  });
  return score;
}
