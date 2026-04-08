(() => {
  if (window.__goRockCampUfretBridgeLoaded) return;
  window.__goRockCampUfretBridgeLoaded = true;

  window.postMessage({ source: 'ufret-extension-bridge', action: 'bridge-ready' }, '*');

  window.addEventListener('message', (event) => {
    const message = event.data || {};
    if (message.source !== 'grockcamp-db-tool') return;
    if (message.action !== 'extractUfretPreviewFromReferenceUrl') return;

    const requestId = String(message.requestId || '');
    chrome.runtime.sendMessage({
      action: 'extractUfretPreviewFromReferenceUrl',
      referenceUrl: message.referenceUrl || ''
    }, (response) => {
      const runtimeError = chrome.runtime.lastError;
      if (runtimeError) {
        window.postMessage({
          source: 'ufret-extension-bridge',
          requestId,
          ok: false,
          error: runtimeError.message || 'runtime message failed'
        }, '*');
        return;
      }
      if (!response) {
        window.postMessage({
          source: 'ufret-extension-bridge',
          requestId,
          ok: false,
          error: 'empty response from extension background'
        }, '*');
        return;
      }
      window.postMessage({
        source: 'ufret-extension-bridge',
        requestId,
        ok: !!response.ok,
        payload: response.ok ? response : undefined,
        error: response.ok ? undefined : (response.message || 'failed')
      }, '*');
    });
  });
})();
