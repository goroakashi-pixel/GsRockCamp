(() => {
  if (window.__goRockCampUfretBridgeLoaded) return;
  window.__goRockCampUfretBridgeLoaded = true;

  postBridgeMessage_({ action: 'bridge-ready', detail: 'content script loaded' });

  window.addEventListener('message', (event) => {
    const message = event.data || {};
    if (message.source !== 'grockcamp-db-tool') return;
    if (message.action !== 'extractUfretPreviewFromReferenceUrl') return;

    const requestId = String(message.requestId || '');
    postBridgeMessage_({ requestId, action: 'bridge-received', detail: 'request accepted in content script' });
    chrome.runtime.sendMessage({
      action: 'extractUfretPreviewFromReferenceUrl',
      referenceUrl: message.referenceUrl || ''
    }, (response) => {
      const runtimeError = chrome.runtime.lastError;
      if (runtimeError) {
        postBridgeMessage_({ requestId, ok: false, error: runtimeError.message || 'runtime message failed' });
        return;
      }
      if (!response) {
        postBridgeMessage_({ requestId, ok: false, error: 'empty response from extension background' });
        return;
      }
      postBridgeMessage_({
        requestId,
        ok: !!response.ok,
        payload: response.ok ? response : undefined,
        error: response.ok ? undefined : (response.message || 'failed')
      });
    });
  });

  function postBridgeMessage_(payload) {
    window.postMessage(Object.assign({ source: 'ufret-extension-bridge' }, payload || {}), '*');
  }
})();
