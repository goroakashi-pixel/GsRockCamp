const SPREADSHEET_ID = '1D6d0iNhMdZn8I0Jj-m1tAT0blZIgec4qTHlAJuGo4ms';
const SHEET_NAME = 'DB';
const PART_ORDER = ['intro', 'A', 'B', 'サビ'];
const SAVE_MODE = 'EMPTY_ONLY'; // 'EMPTY_ONLY' | 'FORCE'
const BAR_COUNT = 8;
const APP_VERSION = '1.2.1';
const OPENAI_MODEL = 'gpt-5-mini';
const ENABLE_RULE_BASED_FALLBACK = false;
const AI_JSON_RETRY_MAX = 2;
const REQUIRED_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/script.external_request'
];

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

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.appVersion = APP_VERSION;
  return template.evaluate()
    .setTitle('GoRockCamp DB投入支援ツール β v' + APP_VERSION)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function suggestOriginalChords(baseId) {
  try {
    var request = normalizeSuggestRequest_(baseId);
    var normalizedBaseId = normalizeBaseId_(request.baseId);
    var bundle = fetchSongBundle_(normalizedBaseId);
    var response = buildSongResponse_(bundle, {
      ok: true,
      action: 'suggestOriginalChords',
      stage: 'ai_chord_drafted'
    });
    response.metadata = response.metadata || {};
    response.metadata.referenceUrl = request.referenceUrl || '';
    var research = researchYouTubeOnly_(bundle);
    applyResearchMetadataToResponse_(response, research);
    response.logs = (response.logs || []).concat((research.logs || []).map(function(line){ return '[ai][youtube] ' + line; }));

    response.logs = (response.logs || []).concat(['[ai][ufret] reference url input: ' + (request.referenceUrl || '(none)')]);
    var ufretRaw = normalizeUfretRawPayload_(request.ufretRawLines, request.ufretRawText);
    if (ufretRaw.text) {
      response.logs.push('[ai][ufret] using cached raw data from step-3');
    }

    if (!ufretRaw.text) {
      response.stage = 'metadata_only';
      response.metadata.draftNotes = (response.metadata.draftNotes || []);
      if (response.metadata.draftNotes.indexOf('U-FRET内で有効なコードソースを取得できませんでした。') < 0) {
        response.metadata.draftNotes.push('U-FRET内で有効なコードソースを取得できませんでした。');
      }
      response.metadata.ufretRawLines = [];
      response.metadata.ufretRawText = '';
      response.metadata.ufretStatus = 'failed';
      response.metadata.ufretSourceUrl = sanitizeUrlInput_(request.referenceUrl);
      response.logs.push('[ai][ufret] skip openai: no chord source');
      response.message = 'U-FRET 参照URLから original_Chord を取得できませんでした。URLを確認してください。';
      return response;
    }

    response.metadata.ufretRawLines = ufretRaw.lines.slice(0);
    response.metadata.ufretRawText = ufretRaw.text;
    response.metadata.ufretStatus = 'ready';
    response.metadata.ufretSourceUrl = sanitizeUrlInput_(request.referenceUrl);
    response.metadata.draftNotes = ufretRaw.previewLines.slice(0, 120);

    var aiConfig = getOpenAiConfig_(true);
    response.logs.push('AI precheck: OPENAI_API_KEY ' + (aiConfig.configured ? 'configured' : 'missing'));
    if (!aiConfig.configured) {
      response.stage = 'metadata_only';
      response.logs.push('[ai][ufret] skip openai: OPENAI_API_KEY missing');
      response.message = 'U-FRET 生データは取得済みです。OPENAI_API_KEY 設定後に「AIでoriginal_Chord取得」を実行してください。';
      return response;
    }
    response.logs.push('[ai][ufret] openai call start');
    var draft = requestAiOriginalChordDraftFromUfretRaw_(bundle, aiConfig, ufretRaw.text);
    response.logs = response.logs.concat(draft.logs || []);
    applyAiDraftToResponse_(response, draft);
    response.metadata.draftNotes = (response.metadata.draftNotes || []).concat(draft.notes || []);
    var appliedCount = countAppliedAiDraftCellsInResponse_(response.parts);
    if (appliedCount <= 0) {
      response.stage = 'metadata_only';
      response.message = '参照URLのコード抽出は成功しましたが、下書き反映に失敗しました。';
      response.logs.push('[ai][ufret] UI apply failure');
      return response;
    }
    response.message = '参照URL(U-FRET)から original_Chord 下書きを取得し、' + appliedCount + 'セルを反映しました。';
    response.logs.push('[ai][ufret] original chord draft created: ' + appliedCount + ' cells');
    return response;
  } catch (error) {
    return buildErrorResponse_(error, { action: 'suggestOriginalChords', baseId: baseId });
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
