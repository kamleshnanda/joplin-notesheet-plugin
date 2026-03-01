import joplin from 'api';
import { MenuItemLocation, ToolbarButtonLocation, ContentScriptType, ModelType } from 'api/types';

/**
 * Joplin Univer Plugin
 * Embeds interactive Univer spreadsheets within Joplin notes
 * VERSION: 2.4.0 - Dialog-based editing with userData storage, RTE support, formulas, row/col manipulation
 */
joplin.plugins.register({
	onStart: async function() {
		try {
			console.info('=== UNIVER PLUGIN v2.4.0 STARTING ===');
			console.info('Joplin Univer Sheets plugin starting...');

			// Create dialog for spreadsheet editing
			const dialogHandle = await joplin.views.dialogs.create('univerDialog');
			await joplin.views.dialogs.setFitToContent(dialogHandle, false); // Use full size
			
			// Set dialog buttons
			await joplin.views.dialogs.setButtons(dialogHandle, [
				{
					id: 'save',
					title: 'Save & Close'
				},
				{
					id: 'cancel',
					title: 'Cancel'
				}
			]);

			// Register content script for rendering Univer links in Markdown viewer
			console.info('Registering Markdown content script...');
			await joplin.contentScripts.register(
				ContentScriptType.MarkdownItPlugin,
				'univerSheet',
				'./contentScript.js'
			);
			console.info('=== MARKDOWN CONTENT SCRIPT REGISTERED SUCCESSFULLY ===');
			
			// Register content script for handling clicks in Rich Text Editor (CodeMirror)
			console.info('Registering CodeMirror content script...');
			await joplin.contentScripts.register(
				ContentScriptType.CodeMirrorPlugin,
				'univerSheetCodeMirror',
				'./contentScriptTinyMCE.js'
			);
			console.info('=== CODEMIRROR CONTENT SCRIPT REGISTERED SUCCESSFULLY ===');

			// Set up message handler for Markdown viewer content script
			await joplin.contentScripts.onMessage('univerSheet', async (sheetId: string) => {
				console.info('Received message from Markdown content script:', sheetId);
				
				// The message is just the sheet ID string
				if (sheetId && typeof sheetId === 'string' && sheetId.startsWith('sheet-')) {
					console.info('Opening spreadsheet from Markdown viewer click:', sheetId);
					await joplin.commands.execute('openUniverSheet', sheetId);
				} else {
					console.warn('Invalid sheet ID received from Markdown viewer:', sheetId);
				}
			});
			
			// Set up message handler for CodeMirror content script
			await joplin.contentScripts.onMessage('univerSheetCodeMirror', async (sheetId: string) => {
				console.info('Received message from CodeMirror content script:', sheetId);
				
				// The message is just the sheet ID string
				if (sheetId && typeof sheetId === 'string' && sheetId.startsWith('sheet-')) {
					console.info('Opening spreadsheet from CodeMirror click:', sheetId);
					await joplin.commands.execute('openUniverSheet', sheetId);
				} else {
					console.warn('Invalid sheet ID received from CodeMirror:', sheetId);
				}
			});

			// Register command to insert Univer spreadsheet
			await joplin.commands.register({
				name: 'insertUniverSheet',
				label: 'Insert Univer Spreadsheet',
				iconName: 'fas fa-table',
				execute: async () => {
					try {
						// Check if a note is selected
						const note = await joplin.workspace.selectedNote();
						if (!note) {
							await joplin.views.dialogs.showMessageBox('Please select a note first');
							return;
						}

						// Generate unique sheet ID
						const sheetId = `sheet-${Date.now()}`;
						
						// Default empty spreadsheet data
						const defaultData = {
							id: sheetId,
							name: 'Spreadsheet',
							sheetOrder: ['sheet-01'],
							appVersion: '0.1.0',
							locale: 'en-US',
							styles: {},
							sheets: {
								'sheet-01': {
									id: 'sheet-01',
									name: 'Sheet1',
									cellData: {
										0: {
											0: { v: 'Column A' },
											1: { v: 'Column B' },
											2: { v: 'Column C' },
										},
									},
									rowCount: 100,
									columnCount: 26,
									defaultColumnWidth: 93,
									defaultRowHeight: 27,
								},
							},
						};

						// Store spreadsheet data in note's userData
						console.info('Storing spreadsheet data in userData with ID:', sheetId);
						await joplin.data.userDataSet(ModelType.Note, note.id, `univer_${sheetId}`, defaultData);
						
						// Create markdown link with custom protocol
						// Use univer-sheet:// instead of joplin://x-callback-url/ to avoid
						// Joplin's internal callback URL handler intercepting the link
						const markdown = `\n[📊 Spreadsheet: ${defaultData.name}](univer-sheet://${sheetId})\n`;
						
						console.info('Inserting markdown link:', markdown);
						
						// Insert at cursor position
						await joplin.commands.execute('insertText', markdown);
						
						// Show success message
						await joplin.views.dialogs.showMessageBox(
							'✅ Univer spreadsheet inserted!\n\n' +
							'Click the link to open the spreadsheet editor.\n' +
							'Your data is safely stored and will sync across devices.'
						);
						
						console.info('Univer spreadsheet inserted successfully');
					} catch (error) {
						console.error('Failed to insert spreadsheet:', error);
						await joplin.views.dialogs.showMessageBox(`Failed to insert spreadsheet: ${error.message}`);
					}
				},
			});

			// Register command to open spreadsheet editor
			await joplin.commands.register({
				name: 'openUniverSheet',
				label: 'Open Univer Spreadsheet',
				execute: async (sheetIdParam?: string) => {
					try {
						let sheetId = sheetIdParam;
						
						// If no sheet ID provided, try to extract from selected text
						if (!sheetId) {
							const selectedText = await joplin.commands.execute('selectedText');
							if (selectedText) {
								// Try to extract sheet ID from various formats:
								// 1. "sheet-1234567890"
								// 2. "ID: sheet-1234567890"
								// 3. "ID: `sheet-1234567890`"
								const match = selectedText.match(/sheet-\d+/);
								if (match) {
									sheetId = match[0];
								}
							}
						}
						
						if (!sheetId) {
							await joplin.views.dialogs.showMessageBox(
								'No spreadsheet ID found.\n\n' +
								'Please select the sheet ID text (e.g., sheet-1234567890) and try again.'
							);
							return;
						}
						
						console.info('Opening spreadsheet editor for ID:', sheetId);
						
						// Get current note
						const note = await joplin.workspace.selectedNote();
						if (!note) {
							await joplin.views.dialogs.showMessageBox('No note selected');
							return;
						}

						// Load spreadsheet data from userData
						const data: any = await joplin.data.userDataGet(ModelType.Note, note.id, `univer_${sheetId}`);
						
						if (!data) {
							await joplin.views.dialogs.showMessageBox('Spreadsheet data not found. It may have been deleted or the note may have changed.');
							return;
						}

						// Ensure data is a proper object for serialization
						let dataObj: any = data;
						if (typeof dataObj === 'string') {
							try { dataObj = JSON.parse(dataObj); } catch(e) { /* use as-is */ }
						}
						
						// Validate data structure
						if (!dataObj || typeof dataObj !== 'object' || !dataObj.sheets) {
							console.error('Invalid spreadsheet data structure:', dataObj);
							await joplin.views.dialogs.showMessageBox(
								'Spreadsheet data is corrupted or in an unexpected format.\n\n' +
								'Data type: ' + typeof dataObj
							);
							return;
						}
						const dataName = (dataObj && dataObj.name) || 'Spreadsheet';
						const sheetCount = (dataObj && dataObj.sheetOrder) ? dataObj.sheetOrder.length : 1;

						// Create dialog HTML — NO inline <script> tags (blocked by CSP).
						// Data is passed via a hidden div; JS is loaded via addScript.
						const html = `
							<style>
								body {
									margin: 0;
									padding: 20px;
									font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif;
									background: #f6f8fa;
								}
								.header {
									background: white;
									padding: 16px;
									border-radius: 6px;
									margin-bottom: 16px;
									box-shadow: 0 1px 3px rgba(0,0,0,0.1);
								}
								.header h2 { margin: 0 0 8px 0; color: #24292f; }
								.header p { margin: 0; color: #57606a; font-size: 14px; }
								.spreadsheet-container {
									background: white;
									border-radius: 6px;
									padding: 16px;
									box-shadow: 0 1px 3px rgba(0,0,0,0.1);
									min-height: 400px;
									overflow: auto;
								}
								table { width: 100%; border-collapse: collapse; font-size: 13px; }
								th, td { border: 1px solid #d0d7de; padding: 6px 10px; text-align: left; }
								th { background: #f6f8fa; font-weight: 600; color: #57606a; }
								input {
									width: 100%; border: none; background: transparent;
									font-family: inherit; font-size: inherit; padding: 0;
									box-sizing: border-box;
								}
								input:focus { outline: 2px solid #0969da; outline-offset: -2px; }
								.formula-cell { background: #f0f7ff; }
								.formula-cell input { color: #0550ae; }
								.info {
									margin-top: 16px; padding: 12px;
									background: #ddf4ff; border-radius: 6px;
									color: #0969da; font-size: 13px;
								}
								#spreadsheet-data { display: none; }
							</style>
							<div class="header">
								<h2>📊 ${dataName}</h2>
								<p>Sheet ID: ${sheetId} • ${sheetCount} sheet(s)</p>
							</div>
							<div class="spreadsheet-container">
								<form name="spreadsheet">
									<div id="spreadsheet-editor"><p style="color:#888">Loading spreadsheet...</p></div>
								</form>
							</div>
							<div class="info">
								💡 Edit cells directly. Use formulas like =SUM(A1:A5), =AVERAGE, =IF. Click "Save &amp; Close" to persist.
							</div>
							<div id="spreadsheet-data">${JSON.stringify(dataObj).replace(/</g, '\\u003c')}</div>
						`;
						
						await joplin.views.dialogs.setHtml(dialogHandle, html);
						await joplin.views.dialogs.addScript(dialogHandle, './dialogScript.js');
						
						// Open dialog and wait for user action
						const result = await joplin.views.dialogs.open(dialogHandle);
						
						console.info('Dialog closed with result:', result);
						
						// If user clicked Save, persist the edited data back to userData
						if (result.id === 'save') {
							const formData = result.formData;
							
							let updatedData: any = dataObj;
							const firstSheetId = (updatedData.sheetOrder && updatedData.sheetOrder[0]) || Object.keys(updatedData.sheets)[0];
							
							if (formData && formData.spreadsheet) {
								// Merge form data with in-memory data.
								// The dialog stores formulas in data-formula attributes and
								// displays computed values in the input. The in-memory
								// spreadsheetData (maintained by dialogScript.js) has the
								// authoritative cell data including formula cells { f, v }.
								// Form inputs only capture display values, so we prefer
								// the in-memory cellData which dialogScript.js keeps in sync.
								// However, as a fallback, we also process form data for cells
								// that may not have been tracked in-memory.
								const existingCellData = updatedData.sheets[firstSheetId].cellData || {};
								const newCellData: any = {};
								
								// Start with in-memory data (preserves formulas)
								for (const [rowKey, rowData] of Object.entries(existingCellData)) {
									const row = parseInt(rowKey);
									if (!isNaN(row)) {
										newCellData[row] = {};
										for (const [colKey, cellVal] of Object.entries(rowData as any)) {
											const col = parseInt(colKey);
											if (!isNaN(col) && cellVal) {
												newCellData[row][col] = cellVal;
											}
										}
									}
								}
								
								// Overlay form data for non-formula cells (plain values)
								for (const [key, value] of Object.entries(formData.spreadsheet)) {
									const match = key.match(/^cell-(\d+)-(\d+)$/);
									if (match) {
										const row = parseInt(match[1]);
										const col = parseInt(match[2]);
										if (!newCellData[row]) newCellData[row] = {};
										// Only overwrite if the cell is NOT a formula cell
										const existing = newCellData[row][col];
										if (!existing || !existing.f) {
											if (value !== '') {
												newCellData[row][col] = { v: value };
											} else if (!existing || !existing.f) {
												delete newCellData[row][col];
											}
										}
									}
								}
								updatedData.sheets[firstSheetId].cellData = newCellData;
							}
							
							// Validate JSON size before saving (5 MB limit)
							const jsonStr = JSON.stringify(updatedData);
							const MAX_JSON_SIZE = 5 * 1024 * 1024;
							if (jsonStr.length > MAX_JSON_SIZE) {
								await joplin.views.dialogs.showMessageBox(
									'Spreadsheet data is too large to save (' + 
									Math.round(jsonStr.length / 1024) + ' KB).\n\n' +
									'Maximum allowed size is 5 MB. Please reduce the amount of data.'
								);
								return;
							}
							
							await joplin.data.userDataSet(ModelType.Note, note.id, `univer_${sheetId}`, updatedData);
							console.info('Spreadsheet data saved for:', sheetId);
						}
						
					} catch (error) {
						console.error('Failed to open spreadsheet editor:', error);
						await joplin.views.dialogs.showMessageBox(`Failed to open editor: ${error.message}`);
					}
				},
			});

			// Add toolbar button
			await joplin.views.toolbarButtons.create(
				'univerSheetButton',
				'insertUniverSheet',
				ToolbarButtonLocation.EditorToolbar
			);

			// Add menu item with keyboard shortcut
			await joplin.views.menuItems.create(
				'univerSheetMenu',
				'insertUniverSheet',
				MenuItemLocation.Tools,
				{ accelerator: 'CmdOrCtrl+Shift+U' }
			);

			console.info('Univer Sheets plugin registration complete');
		} catch (error) {
			console.error('Failed to initialize Univer plugin:', error);
		}
	},
});
