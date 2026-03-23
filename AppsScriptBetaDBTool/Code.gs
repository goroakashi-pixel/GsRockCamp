const SPREADSHEET_ID = '1D6d0iNhMdZn8I0Jj-m1tAT0blZIgec4qTHlAJuGo4ms';
const SHEET_NAME = 'DB';
const PART_ORDER = ['intro', 'A', 'B', 'サビ'];
const SAVE_MODE = 'EMPTY_ONLY'; // 'EMPTY_ONLY' | 'FORCE'
const BAR_COUNT = 8;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GoRockCamp DB投入支援ツール β')
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
    return buildSongResponse_(bundle, {
      ok: true,
      message: '読み込み成功: ID ' + normalizedBaseId + '〜' + (normalizedBaseId + 3)
    });
  } catch (error) {
    return buildErrorResponse_(error, {
      action: 'loadSong',
      baseId: baseId
    });
  }
}

function saveSong(payload) {
  try {
    validatePayload_(payload);

    var normalizedBaseId = normalizeBaseId_(payload.baseId);
    var bundle = fetchSongBundle_(normalizedBaseId);
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

      var theory = recalcTheory_(rebuiltBars, targetRow.originalKey, targetRow.quizKey);
      var rowUpdate = buildRowUpdate_(bundle, targetRow, rebuiltBars, theory);
      updates.push(rowUpdate);
      skipped = skipped.concat(rowUpdate.skipped);
    });

    flushUpdates_(bundle.sheet, bundle.headers, updates);

    var reloadedBundle = fetchSongBundle_(normalizedBaseId);
    var response = buildSongResponse_(reloadedBundle, {
      ok: true,
      message: '保存成功: 再読込済み',
      saveMode: SAVE_MODE,
      skipped: skipped
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
    return buildErrorResponse_(error, {
      action: 'saveSong',
      baseId: payload && payload.baseId
    });
  }
}

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('payload が不正です。');
  }

  normalizeBaseId_(payload.baseId);

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

function buildSongResponse_(bundle, extra) {
  var rows = bundle.rows;
  var first = rows[0];
  var response = {
    ok: true,
    action: 'song',
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
      quizKey: first.quizKey
    },
    parts: PART_ORDER.map(function(partKey) {
      var row = bundle.partMap[partKey];
      return serializePartRow_(row);
    }),
    logs: buildLoadLogs_(bundle)
  };

  if (extra) {
    Object.keys(extra).forEach(function(key) {
      response[key] = extra[key];
    });
  }

  return response;
}

function recalcTheory_(originalChords, originalKey, quizKey) {
  var sourceKey = normalizeKeyName_(originalKey || quizKey || 'C');
  var targetKey = normalizeKeyName_(quizKey || originalKey || 'C');
  var shift = getSemitoneShift_(sourceKey, targetKey);

  return {
    changeChord: originalChords.map(function(cell) {
      return transposeCell_(cell, shift);
    }),
    degree: originalChords.map(function(cell) {
      return buildTheoryCell_(cell, function(chord) {
        return getDegreeText_(chord, targetKey);
      });
    }),
    functionText: originalChords.map(function(cell) {
      return buildTheoryCell_(cell, function(chord) {
        return getFunctionText_(chord, targetKey);
      });
    }),
    nonDiatonic: originalChords.map(function(cell) {
      return buildNonDiatonicCell_(cell, targetKey);
    })
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
      rowMap[rowId] = {
        rowNumber: r + 1,
        values: values[r]
      };
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

  return {
    baseId: baseId,
    sheet: sheet,
    headers: headers,
    headerMap: headerMap,
    rows: rows,
    partMap: partMap
  };
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
        changeChord: row.changeChords[index] || '',
        degree: row.degrees[index] || '',
        functionText: row.functions[index] || '',
        nonDiatonic: row.nonDiatonic[index] || ''
      };
    })
  };
}

function buildLoadLogs_(bundle) {
  return [
    'Spreadsheet 接続: OK (' + SPREADSHEET_ID + ')',
    'Sheet 確認: ' + SHEET_NAME,
    '取得ID: ' + bundle.baseId + '〜' + (bundle.baseId + 3),
    '取得行: ' + bundle.rows.map(function(row) { return row.rowNumber; }).join(', ')
  ];
}

function buildErrorResponse_(error, context) {
  return {
    ok: false,
    action: context && context.action ? context.action : 'error',
    baseId: context && context.baseId ? String(context.baseId) : '',
    connection: {
      connected: false,
      spreadsheetId: SPREADSHEET_ID,
      sheetName: SHEET_NAME
    },
    error: error && error.message ? error.message : String(error),
    logs: ['ERROR: ' + (error && error.message ? error.message : String(error))]
  };
}

function buildRowUpdate_(bundle, row, originalChords, theory) {
  var updates = {};
  var skipped = [];
  var updatedKeys = [];

  for (var i = 1; i <= BAR_COUNT; i += 1) {
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'original_Chord_' + i, originalChords[i - 1], true);
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'change_Chord_' + i, theory.changeChord[i - 1], false);
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'degree_' + i, theory.degree[i - 1], false);
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'Function_' + i, theory.functionText[i - 1], false);
    upsertCell_(bundle.headers, row, updates, skipped, updatedKeys, 'non_diatonic_' + i, theory.nonDiatonic[i - 1], false);
  }

  return {
    rowNumber: row.rowNumber,
    part: normalizePartKey_(row.part),
    updates: updates,
    skipped: skipped,
    updatedKeys: updatedKeys
  };
}

function flushUpdates_(sheet, headers, rowUpdates) {
  rowUpdates.forEach(function(rowUpdate) {
    var keys = Object.keys(rowUpdate.updates);
    if (!keys.length) {
      return;
    }
    keys.forEach(function(headerName) {
      var column = headers.indexOf(headerName);
      if (column < 0) {
        throw new Error('更新対象の列が見つかりません: ' + headerName);
      }
      sheet.getRange(rowUpdate.rowNumber, column + 1).setValue(rowUpdate.updates[headerName]);
    });
  });
}

function upsertCell_(headers, row, updates, skipped, updatedKeys, headerName, nextValue, protectOriginal) {
  if (headerName === 'File') {
    return;
  }
  var column = headers.indexOf(headerName);
  if (column < 0) {
    return;
  }
  var currentValue = String(getCellByHeader_(row, headers, headerName) || '');
  var nextText = nextValue == null ? '' : String(nextValue);

  if (protectOriginal && SAVE_MODE === 'EMPTY_ONLY' && currentValue && currentValue !== nextText) {
    skipped.push(headerName + ' は既存値ありのため未更新');
    return;
  }

  if (currentValue === nextText) {
    return;
  }

  updates[headerName] = nextText;
  updatedKeys.push(headerName);
}

function getCellByHeader_(row, headers, headerName) {
  var index = headers.indexOf(headerName);
  if (index < 0) {
    return '';
  }
  return row.sheetValues && row.sheetValues[index] != null ? row.sheetValues[index] : '';
}

function validateBarPayload_(bar, part, barNumber) {
  if (!bar || typeof bar !== 'object') {
    throw new Error(part + ' ' + barNumber + '小節 のデータが不正です。');
  }

  var first = sanitizeChordInput_(bar.firstHalf);
  var second = sanitizeChordInput_(bar.secondHalf);

  if (!first && second) {
    throw new Error(part + ' ' + barNumber + '小節 は後半のみ入力できません。');
  }

  [first, second].forEach(function(chordText) {
    if (!chordText) {
      return;
    }
    if (chordText.indexOf('│') >= 0) {
      throw new Error(part + ' ' + barNumber + '小節 は UI が自動で │ を付与します。');
    }
    validateChordSymbol_(chordText, part, barNumber);
  });
}

function buildOriginalChordCell_(bar, part, barNumber) {
  var first = sanitizeChordInput_(bar.firstHalf);
  var second = sanitizeChordInput_(bar.secondHalf);
  validateBarPayload_(bar, part, barNumber);
  if (first && second) {
    return first + '│' + second;
  }
  return first;
}

function validateChordSymbol_(chordText, part, barNumber) {
  var slashCount = (String(chordText).match(/\//g) || []).length;
  if (slashCount > 1) {
    throw new Error(part + ' ' + barNumber + '小節 に不正な / 連続表記があります: ' + chordText);
  }

  var base = String(chordText).split('/')[0];
  if (!/^([A-G](?:#|b|♭)?)(?:[0-9A-Za-z+()#♭b\-susdimaugmajMadd]*)$/.test(base)) {
    throw new Error(part + ' ' + barNumber + '小節 のコード形式が不正です: ' + chordText);
  }

  if (slashCount === 1) {
    var bass = String(chordText).split('/')[1] || '';
    if (!/^[A-G](?:#|b|♭)?$/.test(bass)) {
      throw new Error(part + ' ' + barNumber + '小節 のオンコード表記が不正です: ' + chordText);
    }
  }
}

function splitChordCell_(cellText) {
  var text = String(cellText || '').trim();
  if (!text) {
    return {
      first: '',
      second: '',
      aiSuggestion: []
    };
  }

  var parts = text.split('│').map(function(item) {
    return item.trim();
  }).filter(Boolean);

  return {
    first: parts[0] || '',
    second: parts[1] || '',
    aiSuggestion: parts
  };
}

function buildTheoryCell_(cellText, mapper) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  if (!parts.length) {
    return '';
  }
  return parts.map(mapper).join('│');
}

function buildNonDiatonicCell_(cellText, keyName) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  if (!parts.length) {
    return '';
  }
  var hasNonDiatonic = parts.some(function(chord) {
    return !isChordDiatonic_(chord, keyName);
  });
  return hasNonDiatonic ? 'TRUE' : 'FALSE';
}

function transposeCell_(cellText, shift) {
  var parts = splitChordCell_(cellText).aiSuggestion;
  if (!parts.length) {
    return '';
  }
  return parts.map(function(chord) {
    return transposeChordSymbol_(chord, shift);
  }).join('│');
}

function transposeChordSymbol_(chordText, shift) {
  var parsed = parseChordSymbol_(chordText);
  var root = transposePitchName_(parsed.root, shift);
  var bass = parsed.bass ? transposePitchName_(parsed.bass, shift) : '';
  return root + parsed.quality + (bass ? '/' + bass : '');
}

function getDegreeText_(chordText, keyName) {
  var parsed = parseChordSymbol_(chordText);
  var roman = getRomanDegree_(parsed.root, keyName);
  return roman + inferDegreeSuffix_(parsed.quality);
}

function getFunctionText_(chordText, keyName) {
  var parsed = parseChordSymbol_(chordText);
  var roman = getRomanDegree_(parsed.root, keyName);
  if (/Ⅴ/.test(roman) || /Ⅶ/.test(roman)) {
    return isChordDiatonic_(chordText, keyName) ? 'D' : 'SecDom';
  }
  if (/Ⅱ|Ⅳ/.test(roman)) {
    return isChordDiatonic_(chordText, keyName) ? 'SD' : 'Borrowed';
  }
  if (/Ⅰ|Ⅵ|Ⅲ/.test(roman)) {
    return isChordDiatonic_(chordText, keyName) ? 'T' : 'Borrowed';
  }
  return 'Borrowed';
}

function isChordDiatonic_(chordText, keyName) {
  var scaleSet = getMajorScaleSet_(keyName);
  var chordNotes = buildChordPitchClasses_(chordText);
  return chordNotes.every(function(note) {
    return scaleSet[note];
  });
}

function buildChordPitchClasses_(chordText) {
  var parsed = parseChordSymbol_(chordText);
  var rootIndex = noteNameToIndex_(parsed.root);
  if (rootIndex < 0) {
    return [];
  }

  var intervals = [0, 4, 7];
  if (/^m(?!aj)/.test(parsed.quality)) {
    intervals = [0, 3, 7];
  }
  if (/dim/.test(parsed.quality)) {
    intervals = [0, 3, 6];
  }
  if (/aug/.test(parsed.quality)) {
    intervals = [0, 4, 8];
  }
  if (/sus4/.test(parsed.quality)) {
    intervals = [0, 5, 7];
  }

  if (/6/.test(parsed.quality)) {
    intervals.push(9);
  }
  if (/maj7|M7|mM7/.test(parsed.quality)) {
    intervals.push(11);
  } else if (/7/.test(parsed.quality)) {
    intervals.push(10);
  }
  if (/9|add9/.test(parsed.quality)) {
    intervals.push(14);
  }

  if (parsed.bass) {
    intervals.push(noteNameToIndex_(parsed.bass) - rootIndex);
  }

  var pitchSet = {};
  intervals.forEach(function(interval) {
    var normalized = ((rootIndex + interval) % 12 + 12) % 12;
    pitchSet[indexToSharpName_(normalized)] = true;
  });

  return Object.keys(pitchSet);
}

function getMajorScaleSet_(keyName) {
  var root = noteNameToIndex_(keyName);
  var majorScale = [0, 2, 4, 5, 7, 9, 11];
  var set = {};
  majorScale.forEach(function(interval) {
    set[indexToSharpName_((root + interval) % 12)] = true;
  });
  return set;
}

function getRomanDegree_(rootName, keyName) {
  var diff = ((noteNameToIndex_(rootName) - noteNameToIndex_(keyName)) % 12 + 12) % 12;
  var map = {
    0: 'Ⅰ',
    1: '♭Ⅱ',
    2: 'Ⅱ',
    3: '♭Ⅲ',
    4: 'Ⅲ',
    5: 'Ⅳ',
    6: '♭Ⅴ',
    7: 'Ⅴ',
    8: '♭Ⅵ',
    9: 'Ⅵ',
    10: '♭Ⅶ',
    11: 'Ⅶ'
  };
  return map[diff] || '?';
}

function inferDegreeSuffix_(qualityText) {
  var quality = String(qualityText || '');
  if (!quality) {
    return '';
  }
  if (/m7-5/.test(quality)) {
    return 'm7-5';
  }
  if (/maj7|M7/.test(quality)) {
    return 'maj7';
  }
  if (/m6/.test(quality)) {
    return 'm6';
  }
  if (/6/.test(quality)) {
    return '6';
  }
  if (/m7/.test(quality)) {
    return 'm7';
  }
  if (/m(?!aj)/.test(quality)) {
    return 'm';
  }
  if (/dim/.test(quality)) {
    return 'dim';
  }
  if (/aug/.test(quality)) {
    return 'aug';
  }
  if (/sus4/.test(quality)) {
    return 'sus4';
  }
  if (/add9/.test(quality)) {
    return '(add9)';
  }
  if (/9/.test(quality)) {
    return '9';
  }
  if (/7/.test(quality)) {
    return '7';
  }
  return quality;
}

function parseChordSymbol_(chordText) {
  var text = String(chordText || '').trim();
  var halves = text.split('/');
  var main = halves[0] || '';
  var bass = halves[1] || '';
  var match = main.match(/^([A-G](?:#|b|♭)?)(.*)$/);
  if (!match) {
    throw new Error('コードを解釈できません: ' + chordText);
  }
  return {
    root: normalizeKeyName_(match[1]),
    quality: match[2] || '',
    bass: bass ? normalizeKeyName_(bass) : ''
  };
}

function getSemitoneShift_(fromKey, toKey) {
  return ((noteNameToIndex_(toKey) - noteNameToIndex_(fromKey)) % 12 + 12) % 12;
}

function transposePitchName_(noteName, shift) {
  var index = noteNameToIndex_(noteName);
  if (index < 0) {
    return noteName;
  }
  return indexToSharpName_((index + shift) % 12);
}

function noteNameToIndex_(noteName) {
  var normalized = normalizeKeyName_(noteName);
  var map = {
    C: 0,
    'B#': 0,
    'C#': 1,
    Db: 1,
    D: 2,
    'D#': 3,
    Eb: 3,
    E: 4,
    Fb: 4,
    'E#': 5,
    F: 5,
    'F#': 6,
    Gb: 6,
    G: 7,
    'G#': 8,
    Ab: 8,
    A: 9,
    'A#': 10,
    Bb: 10,
    B: 11,
    Cb: 11
  };
  return map.hasOwnProperty(normalized) ? map[normalized] : -1;
}

function indexToSharpName_(index) {
  return ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B'][((index % 12) + 12) % 12];
}

function normalizeKeyName_(text) {
  return String(text || '')
    .trim()
    .replace(/♭/g, 'b')
    .replace(/＃/g, '#')
    .replace(/([A-Ga-g])/, function(match) {
      return match.toUpperCase();
    });
}

function normalizePartKey_(partText) {
  var normalized = String(partText || '').trim();
  var map = {
    intro: 'intro',
    イントロ: 'intro',
    A: 'A',
    Aメロ: 'A',
    B: 'B',
    Bメロ: 'B',
    サビ: 'サビ',
    chorus: 'サビ'
  };
  return map.hasOwnProperty(normalized) ? map[normalized] : normalized;
}

function normalizeBaseId_(baseId) {
  var text = normalizeNumericString_(baseId);
  if (!text) {
    throw new Error('先頭IDは数値で入力してください。');
  }
  return Number(text);
}

function normalizeNumericString_(value) {
  var text = String(value == null ? '' : value).trim();
  if (!/^\d+$/.test(text)) {
    return '';
  }
  return text;
}

function sanitizeChordInput_(value) {
  return String(value == null ? '' : value)
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/｜/g, '│');
}

function toBooleanLikeString_(value) {
  var text = String(value == null ? '' : value).trim().toUpperCase();
  if (text === 'TRUE') {
    return 'TRUE';
  }
  if (text === 'FALSE') {
    return 'FALSE';
  }
  return text || '';
}

function buildYoutubeEmbedUrl_(rawValue) {
  var text = String(rawValue || '').trim();
  if (!text) {
    return '';
  }
  var idMatch = text.match(/(?:v=|youtu\.be\/|embed\/|shorts\/)([A-Za-z0-9_-]{11})/);
  var videoId = idMatch ? idMatch[1] : text;
  if (!/^[A-Za-z0-9_-]{11}$/.test(videoId)) {
    return '';
  }
  return 'https://www.youtube.com/embed/' + videoId;
}

function getByAliases_(values, headers, headerMap, aliases) {
  var column = findColumnIndex_(headers, headerMap, aliases);
  return column >= 0 ? values[column] : '';
}

function getBarColumn_(values, headers, headerMap, baseName, index) {
  var aliases = [
    baseName + index,
    baseName + '_' + index,
    baseName + String(index),
    baseName.toLowerCase() + '_' + index,
    baseName.toUpperCase() + '_' + index
  ];
  var column = findColumnIndex_(headers, headerMap, aliases);
  return column >= 0 ? values[column] : '';
}

function buildHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    map[normalizeHeader_(header)] = index;
  });
  return map;
}

function findColumnIndex_(headers, headerMap, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    var key = normalizeHeader_(aliases[i]);
    if (headerMap.hasOwnProperty(key)) {
      return headerMap[key];
    }
  }

  for (var j = 0; j < headers.length; j += 1) {
    var normalizedHeader = normalizeHeader_(headers[j]);
    for (var k = 0; k < aliases.length; k += 1) {
      if (normalizedHeader === normalizeHeader_(aliases[k])) {
        return j;
      }
    }
  }

  return -1;
}

function normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-]/g, '')
    .replace(/[^a-z0-9\u3040-\u30ff\u4e00-\u9fff]/g, '');
}
