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

function buildDraftFromReferenceChords_(bundle, chordResult, youtubeUrl) {
  var chordPool = chordResult.chords || [];
  var originalKey = '';
  if (chordResult.keyHints && chordResult.keyHints.length) originalKey = sanitizeKeyInput_(chordResult.keyHints[0]);
  if (!originalKey) originalKey = inferOriginalKeyFromChordList_(chordPool) || sanitizeKeyInput_(bundle.rows[0].originalKey) || '';
  var draft = {
    originalKey: originalKey,
    quizKey: determineQuizKey_(originalKey),
    youtubeUrl: sanitizeUrlInput_(youtubeUrl),
    notes: ['＜音寧コメ＞参照URLから抽出したU-FRETコードを下書き化しました。'],
    partMap: {}
  };
  PART_ORDER.forEach(function(part, partIndex) {
    draft.partMap[part] = [];
    var bars = buildBarsFromChordPool_(chordPool, partIndex * BAR_COUNT * 2, BAR_COUNT);
    for (var i = 0; i < BAR_COUNT; i += 1) draft.partMap[part].push(bars[i]);
  });
  return draft;
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

function runAiDraftAttemptFromUfretRaw_(bundle, config, ufretRawText, attempt, previousRawText, retryHint) {
  return callOpenAiResponsesApiRaw_(config.apiKey, {
    model: config.model,
    input: buildUfretStructuringPrompt_(bundle, ufretRawText, attempt > 1 ? previousRawText : '', attempt > 1 ? retryHint : ''),
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

function researchYouTubeOnly_(bundle) {
  var first = bundle.rows[0];
  var youtubeSearch = searchYoutubeCandidate_(first.artist, first.title);
  var youtubeUrl = youtubeSearch.candidate ? youtubeSearch.candidate.url : '';
  var logs = ['search start'].concat(youtubeSearch.logs || []);
  if (youtubeUrl) logs.push('candidate found: ' + youtubeUrl);
  else logs.push('no match');
  return {
    ok: true,
    originalKey: '',
    originalKeySource: '',
    youtubeUrl: youtubeUrl,
    youtubeSource: youtubeUrl ? 'web_research' : '',
    sources: youtubeSearch.candidate ? [youtubeSearch.candidate] : [],
    logs: logs,
    notes: youtubeUrl ? ['公開Web調査から YouTube 候補を取得しました。公式動画か確認してください。'] : ['公開Web調査では YouTube 候補を取得できませんでした。検索リンクで確認してください。']
  };
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

function buildUfretStructuringPrompt_(bundle, ufretRawText, previousRawText, retryHint) {
  var first = bundle.rows[0];
  var retrySection = previousRawText ? [
    '前回出力がJSONスキーマ不一致でした。JSON以外を一切含めず、schema準拠オブジェクトのみ返してください。',
    '--- previous output start ---',
    String(previousRawText || '').slice(0, 4000),
    '--- previous output end ---',
    String(retryHint || '')
  ].join('\n') : String(retryHint || '');
  return [
    '以下は U-FRET 由来の [コード]歌詞 生データです。これのみを根拠に original_Chord 下書きを構造化してください。',
    '外部web検索は禁止。推測で埋めすぎない。',
    'Artist: ' + first.artist,
    'Title: ' + first.title,
    'Current DB original_key: ' + (first.originalKey || ''),
    '出力制約:',
    '- part は intro / A / B / サビ',
    '- bars は各 part 8件',
    '- bars は original_key（原曲キー）で返す',
    '- secondHalf only 禁止',
    '- 1小節2コード時だけ firstHalf/secondHalf を使用',
    '- 各partに最低1小節以上は firstHalf を埋めること（全part空欄は不可）',
    '- 不明な箇所は空文字',
    '--- ufret_raw start ---',
    String(ufretRawText || '').slice(0, 24000),
    '--- ufret_raw end ---',
    retrySection
  ].join('\n');
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

function sanitizeAiChordDraftValue_(value) { return sanitizeChordInput_(String(value == null ? '' : value).replace(/\|/g, '│')); }

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

function prefixLogs_(logs, prefix) {
  return (logs || []).map(function(line) {
    var text = String(line || '');
    if (!text) return text;
    if (text.indexOf(prefix) === 0) return text;
    return prefix + ' ' + text;
  });
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

function summarizeRawText_(text, maxLength) {
  var normalized = String(text || '').replace(/\s+/g, ' ').trim();
  if (!normalized) return '(empty)';
  return normalized.slice(0, maxLength || 300);
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

function extractJsonErrorPosition_(message) {
  var match = String(message || '').match(/position\s+(\d+)/i);
  return match ? Number(match[1]) : -1;
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

function researchPublicSongData_(bundle, options) {
  options = options || {};
  var first = bundle.rows[0];
  var query = [first.artist || '', first.title || ''].join(' ').trim();
  if (!query) throw new Error('artist/title が空のため Web 調査できません。');

  var logs = ['[ai] research start: ' + query];
  var notes = [];
  var sources = [];
  logs.push('[ai][youtube] search start');
  var youtubeSearch = searchYoutubeCandidate_(first.artist, first.title);
  var youtubeResult = youtubeSearch.candidate;
  logs = logs.concat(prefixLogs_(youtubeSearch.logs || [], '[ai][youtube]'));
  if (youtubeResult) {
    sources.push(youtubeResult);
    logs.push('[ai][youtube] candidate found: ' + youtubeResult.url);
  } else {
    logs.push('[ai][youtube] no match');
  }

  logs.push('[ai][chord] U-FRET search start');
  var chordSearch = findUfretChordCandidates_(first.artist, first.title);
  var chordResults = chordSearch.results || [];
  logs = logs.concat(prefixLogs_(chordSearch.logs || [], '[ai][chord]'));
  chordResults.forEach(function(result) { sources.push(result); });
  if (chordResults.length) logs.push('[ai][chord] U-FRET candidate found: ' + chordResults.length);
  else logs.push('[ai][chord] U-FRET no match');

  var scrapedChordPool = [];
  var acceptedChordPageCount = 0;
  chordResults.slice(0, 5).forEach(function(result) {
    try {
      var html = fetchHtmlText_(result.url);
      var pageCheck = validateUfretChordPage_(result, html, first.originalKey);
      if (!pageCheck.accepted) {
        logs.push('[ai][chord] U-FRET page rejected: ' + pageCheck.reason + ' / ' + result.url);
        return;
      }
      var chords = extractChordCandidatesFromHtml_(html);
      var keyHints = extractOriginalKeyHintsFromHtml_(html);
      result.chords = chords.slice(0, 64);
      result.keyHints = keyHints.slice(0, 8);
      if (result.chords.length) {
        acceptedChordPageCount += 1;
        scrapedChordPool = scrapedChordPool.concat(result.chords);
        logs.push('[ai][chord] U-FRET scrape success: ' + result.url + ' / chords=' + result.chords.length);
      } else {
        logs.push('[ai][chord] U-FRET scrape failed: chords empty / ' + result.url);
      }
      if (result.keyHints.length) logs.push('[ai][chord] key hints: ' + result.keyHints.join(', '));
    } catch (error) {
      logs.push('[ai][chord] U-FRET scrape failed: ' + result.url + ' / ' + error.message);
    }
  });
  if (!acceptedChordPageCount) logs.push('[ai][chord] U-FRET no usable page');

  var originalKey = pickOriginalKeyFromSources_(sources, scrapedChordPool);
  var youtubeUrl = youtubeResult ? youtubeResult.url : '';
  if (originalKey) notes.push('U-FRET から original_key 候補を取得しました。保存前に確認してください。');
  if (youtubeUrl) notes.push('公開Web調査から YouTube 候補を取得しました。公式動画か確認してください。');
  if (!originalKey) notes.push('U-FRET 該当なし、または条件不一致のため original_Chord 下書きは作成できませんでした。');
  if (!youtubeUrl) notes.push('公開Web調査では YouTube 候補を取得できませんでした。検索リンクで確認してください。');
  if (!options.metadataOnly && !getOpenAiConfig_(true).configured) {
    logs.push('[ai] OPENAI_API_KEY 未設定のため original_Chord AI draft は未実行');
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

function sanitizeAiNote_(note) {
  var text = String(note == null ? '' : note).trim();
  text = text.replace(/https?:\/\/\S+/g, '').replace(/\[[^\]]+\]\([^)]+\)/g, '').trim();
  return text.slice(0, 120);
}

function requestAiOriginalChordDraftFromUfretRaw_(bundle, config, ufretRawText) {
  config = config && config.configured ? config : getOpenAiConfig_();
  var rawText = String(ufretRawText || '').trim();
  if (!rawText) throw new Error('U-FRET 生データが空です。');
  var logs = [
    'OpenAI 実行開始（U-FRET生データ優先）',
    'AI draft model: ' + config.model + ' (' + (config.modelSource || 'default') + ')',
    'U-FRET raw length: ' + rawText.length
  ];
  var lastError = null;
  var rawTextForRetry = '';
  var retryHint = '';
  for (var attempt = 1; attempt <= AI_JSON_RETRY_MAX; attempt += 1) {
    try {
      logs.push('AI JSON生成 attempt ' + attempt + '/' + AI_JSON_RETRY_MAX);
      var aiJson = runAiDraftAttemptFromUfretRaw_(bundle, config, rawText, attempt, rawTextForRetry, retryHint);
      rawTextForRetry = extractResponseText_(aiJson);
      var parsed = extractStructuredDraftObject_(aiJson, logs);
      validateAiDraftSchema_(parsed);
      var draft = normalizeAiDraftResponse_(parsed, bundle);
      var keyCheck = checkDraftBarsKeyAgainstOriginal_(draft, bundle);
      if (keyCheck.retry) throw new Error('AI bars appear transposed to Quiz_key');
      var draftCount = countNonEmptyDraftCells_(draft.partMap);
      if (draftCount <= 0) throw new Error('AI returned empty bars');
      logs.push('AI下書き受領');
      logs.push('JSON検証OK');
      logs.push('non-empty draft count: ' + draftCount);
      draft.logs = logs;
      return draft;
    } catch (error) {
      lastError = error;
      var errorType = classifyAiDraftError_(error);
      logs.push(errorType + ' attempt ' + attempt + ': ' + error.message);
      logs.push('AI raw snippet: ' + summarizeRawText_(rawTextForRetry, 800));
      retryHint = /empty bars/i.test(String(error && error.message || ''))
        ? '前回はbarsが空でした。U-FRET生データからintro/A/B/サビを最低1小節以上埋めてください。'
        : '前回はJSON整形に失敗しました。JSON以外を一切含めないでください。';
      if (attempt >= AI_JSON_RETRY_MAX) break;
      logs.push('JSON再試行を実行します。');
    }
  }
  var wrapped = new Error('AI JSON整形失敗: ' + (lastError ? lastError.message : 'unknown'));
  wrapped.draftLogs = logs;
  throw wrapped;
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

