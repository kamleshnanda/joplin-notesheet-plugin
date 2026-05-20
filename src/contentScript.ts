/**
 * Content script for rendering Univer spreadsheet links in Joplin notes.
 * 
 * Intercepts Markdown links with univer-sheet:// protocol and renders them
 * as <span> elements (NOT <a> tags) to prevent Electron from intercepting
 * the custom protocol via openExternal/IPC.
 * 
 * VERSION: 2.2.0
 * 
 * Architecture:
 * - This outer function runs in the PLUGIN process (not the webview)
 * - The `plugin` function modifies markdown-it rendering (plugin process)
 * - The `assets` function loads univerWebviewHandler.js into the NOTE VIEWER WEBVIEW
 * - Click handling happens in the webview where webviewApi.postMessage is available
 * - Communication: webview -> plugin via webviewApi.postMessage(contentScriptId, sheetId)
 * 
 * Rich Text Editor support:
 * - Uses joplin-editable + joplin-source pattern to preserve original markdown
 *   when switching between Markdown and Rich Text editors
 */

export default function(context: any) {
    const contentScriptId = context.contentScriptId;
    console.log('=== UNIVER CONTENT SCRIPT v2.2.0 LOADED ===');
    console.log('Content script ID:', contentScriptId);
    
    return {
        /**
         * Markdown-it plugin: replaces univer-sheet:// links with <span> elements
         * wrapped in joplin-editable containers for Rich Text Editor compatibility.
         */
        plugin: function(markdownIt: any, _options: any) {
            console.log('Univer markdown-it plugin initializing...');
            
            const defaultLinkOpen = markdownIt.renderer.rules.link_open || function(tokens: any, idx: any, options: any, _env: any, self: any) {
                return self.renderToken(tokens, idx, options);
            };
            
            const defaultLinkClose = markdownIt.renderer.rules.link_close || function(tokens: any, idx: any, options: any, _env: any, self: any) {
                return self.renderToken(tokens, idx, options);
            };

            let univerLinkHref = '';
            let univerLinkText = '';
            let isUniverLink = false;

            // Override link_open to detect univer links and suppress the <a> tag
            markdownIt.renderer.rules.link_open = function(tokens: any, idx: any, options: any, _env: any, self: any) {
                const token = tokens[idx];
                const hrefIndex = token.attrIndex('href');
                
                if (hrefIndex >= 0) {
                    const href = token.attrs[hrefIndex][1];
                    
                    const univerMatch = href && (
                        href.match(/^joplin:\/\/x-callback-url\/univer\/(.+)/) ||
                        href.match(/^univer-sheet:\/\/(.+)/)
                    );
                    
                    if (univerMatch) {
                        isUniverLink = true;
                        univerLinkHref = href;
                        univerLinkText = '';
                        
                        // Collect the link text from subsequent tokens
                        for (let i = idx + 1; i < tokens.length; i++) {
                            if (tokens[i].type === 'link_close') break;
                            if (tokens[i].type === 'text') {
                                univerLinkText += tokens[i].content;
                            }
                        }
                        
                        const sheetId = univerMatch[1];
                        const originalMarkdown = '[' + univerLinkText + '](' + univerLinkHref + ')';
                        
                        console.log('Univer link detected! Sheet ID:', sheetId);
                        
                        // Wrap in joplin-editable for Rich Text Editor preservation.
                        // The joplin-source pre tag stores the original markdown so
                        // the RTE can convert back without losing the link.
                        let html = '<div class="joplin-editable">';
                        html += '<pre class="joplin-source" data-joplin-language="univer" '
                            + 'data-joplin-source-open="" '
                            + 'data-joplin-source-close="">'
                            + escapeHtml(originalMarkdown)
                            + '</pre>';
                        html += '<span class="univer-link" '
                            + 'data-sheet-id="' + sheetId + '" '
                            + 'data-content-script-id="' + contentScriptId + '" '
                            + 'role="button" tabindex="0">';
                        
                        return html;
                    }
                }
                
                isUniverLink = false;
                return defaultLinkOpen(tokens, idx, options, _env, self);
            };
            
            markdownIt.renderer.rules.link_close = function(tokens: any, idx: any, options: any, _env: any, self: any) {
                if (isUniverLink) {
                    isUniverLink = false;
                    // Close the span and the joplin-editable div
                    return '</span></div>';
                }
                return defaultLinkClose(tokens, idx, options, _env, self);
            };
            
            function escapeHtml(str: string): string {
                return str
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;');
            }
        },
        
        /**
         * Assets loaded into the note viewer webview.
         * - CSS: styles the univer-link spans and hides joplin-source
         * - JS: univerWebviewHandler.js sets up click handling with webviewApi.postMessage
         */
        assets: function() {
            return [
                {
                    inline: true,
                    text: '.joplin-editable .joplin-source { display: none !important; } .univer-link { display: inline-flex; align-items: center; gap: 8px; padding: 8px 16px; background: #f6f8fa; border: 1px solid #d0d7de; border-radius: 6px; text-decoration: none; color: #24292f; font-weight: 500; cursor: pointer; user-select: none; } .univer-link:hover { background: #e8ecf0 !important; border-color: #afb8c1 !important; } .univer-link:focus { outline: 2px solid #0969da; outline-offset: 2px; } .univer-formula { color: #1a7f37; font-weight: 500; } @media (prefers-color-scheme: dark) { .univer-formula { color: #4ac26b; } }',
                    mime: 'text/css',
                },
                {
                    name: 'univerWebviewHandler.js',
                },
            ];
        },
    };
}
