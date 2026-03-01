/**
 * Webview click handler for Univer spreadsheet links.
 * This file runs inside the Joplin note viewer webview context,
 * where `webviewApi` is available as a global.
 * 
 * It handles clicks on <span class="univer-link"> elements
 * and sends the sheet ID back to the plugin via webviewApi.postMessage.
 */
(function() {
    if (window.__univerClickHandlerInstalled) return;
    window.__univerClickHandlerInstalled = true;
    
    console.log('[Univer] Installing click handler in webview...');
    
    document.addEventListener('click', function(event) {
        var target = event.target;
        // Walk up to find .univer-link span
        while (target && target !== document) {
            if (target.classList && target.classList.contains('univer-link')) {
                break;
            }
            target = target.parentElement;
        }
        
        if (!target || !target.classList || !target.classList.contains('univer-link')) {
            return;
        }
        
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        
        var sheetId = target.getAttribute('data-sheet-id');
        var csId = target.getAttribute('data-content-script-id');
        
        console.log('[Univer] Link clicked! Sheet ID:', sheetId, 'Content Script ID:', csId);
        
        if (sheetId && csId && typeof webviewApi !== 'undefined') {
            webviewApi.postMessage(csId, sheetId).then(function(response) {
                console.log('[Univer] Message sent, response:', response);
            }).catch(function(err) {
                console.error('[Univer] postMessage error:', err);
            });
        } else {
            console.error('[Univer] Missing sheetId, csId, or webviewApi not available');
            console.log('[Univer] webviewApi available:', typeof webviewApi !== 'undefined');
        }
        
        return false;
    }, true); // capture phase to intercept before any other handlers
    
    // Keyboard accessibility: Enter/Space on focused span
    document.addEventListener('keydown', function(event) {
        if (event.key !== 'Enter' && event.key !== ' ') return;
        
        var target = event.target;
        if (target && target.classList && target.classList.contains('univer-link')) {
            event.preventDefault();
            target.click();
        }
    }, true);
    
    console.log('[Univer] Click handler installed in webview');
})();
