import joplin from 'api';
import { MenuItemLocation, ToolbarButtonLocation, ContentScriptType, ModelType } from 'api/types';

const LOG = '[Univer]';
const SHEET_ID_PREFIX = 'sheet-';

// Run a registration step and log success/failure without aborting the rest of onStart.
async function step(name: string, fn: () => Promise<void>): Promise<void> {
	try {
		await fn();
		console.info(LOG, name, 'OK');
	} catch (error) {
		console.error(LOG, name, 'FAILED', error);
	}
}

/**
 * Joplin Univer Plugin
 * Embeds interactive Univer spreadsheets within Joplin notes
 * VERSION: 2.4.0 - Dialog-based editing with userData storage, RTE support, formulas, row/col manipulation
 */
joplin.plugins.register({
	onStart: async function() {
		console.info(LOG, '=== UNIVER PLUGIN v2.4.0 STARTING ===');

		let dialogHandle = '';

		await step('create dialog', async () => {
			dialogHandle = await joplin.views.dialogs.create('univerDialog');
			await joplin.views.dialogs.setFitToContent(dialogHandle, false);
			await joplin.views.dialogs.setButtons(dialogHandle, [
				{ id: 'save', title: 'Save & Close' },
				{ id: 'cancel', title: 'Cancel' },
			]);
			// addScript is one-time at registration; setHtml is per-open.
			await joplin.views.dialogs.addScript(dialogHandle, './dialogScript.js');
		});

		await step('register MarkdownItPlugin content script', async () => {
			await joplin.contentScripts.register(
				ContentScriptType.MarkdownItPlugin,
				'univerSheet',
				'./contentScript.js'
			);
		});

		await step('register CodeMirrorPlugin content script', async () => {
			await joplin.contentScripts.register(
				ContentScriptType.CodeMirrorPlugin,
				'univerSheetCodeMirror',
				'./contentScriptTinyMCE.js'
			);
		});

		const messageSources = [
			{ id: 'univerSheet', label: 'MarkdownIt' },
			{ id: 'univerSheetCodeMirror', label: 'CodeMirror' },
		];
		for (const { id, label } of messageSources) {
			await step(`register ${label} message handler`, async () => {
				await joplin.contentScripts.onMessage(id, async (sheetId: string) => {
					if (typeof sheetId === 'string' && sheetId.startsWith(SHEET_ID_PREFIX)) {
						console.info(LOG, `open from ${label} click:`, sheetId);
						await joplin.commands.execute('openUniverSheet', sheetId);
					} else {
						console.warn(LOG, `invalid sheet id from ${label}:`, sheetId);
					}
				});
			});
		}

		await step('register insertUniverSheet command', async () => {
			await joplin.commands.register({
				name: 'insertUniverSheet',
				label: 'Insert Univer Spreadsheet',
				iconName: 'fas fa-table',
				execute: async () => {
					try {
						const note = await joplin.workspace.selectedNote();
						if (!note) {
							await joplin.views.dialogs.showMessageBox('Please select a note first');
							return;
						}

						const sheetId = `${SHEET_ID_PREFIX}${Date.now()}`;
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

						await joplin.data.userDataSet(ModelType.Note, note.id, `univer_${sheetId}`, defaultData);

						// Use univer-sheet:// instead of joplin://x-callback-url/ to avoid
						// Joplin's internal callback URL handler intercepting the link.
						const markdown = `\n[📊 Spreadsheet: ${defaultData.name}](univer-sheet://${sheetId})\n`;
						await joplin.commands.execute('insertText', markdown);

						await joplin.views.dialogs.showMessageBox(
							'✅ Univer spreadsheet inserted!\n\n' +
							'Click the link to open the spreadsheet editor.\n' +
							'Your data is safely stored and will sync across devices.'
						);
					} catch (error) {
						console.error(LOG, 'insertUniverSheet failed:', error);
						await joplin.views.dialogs.showMessageBox(`Failed to insert spreadsheet: ${error.message}`);
					}
				},
			});
		});

		await step('register openUniverSheet command', async () => {
			await joplin.commands.register({
				name: 'openUniverSheet',
				label: 'Open Univer Spreadsheet',
				execute: async (sheetIdParam?: string) => {
					try {
						let sheetId = sheetIdParam;

						if (!sheetId) {
							const selectedText = await joplin.commands.execute('selectedText');
							if (selectedText) {
								const match = selectedText.match(/sheet-\d+/);
								if (match) sheetId = match[0];
							}
						}

						if (!sheetId) {
							await joplin.views.dialogs.showMessageBox(
								'No spreadsheet ID found.\n\n' +
								'Please select the sheet ID text (e.g., sheet-1234567890) and try again.'
							);
							return;
						}

						const note = await joplin.workspace.selectedNote();
						if (!note) {
							await joplin.views.dialogs.showMessageBox('No note selected');
							return;
						}

						const data: any = await joplin.data.userDataGet(ModelType.Note, note.id, `univer_${sheetId}`);
						if (!data) {
							await joplin.views.dialogs.showMessageBox('Spreadsheet data not found. It may have been deleted or the note may have changed.');
							return;
						}

						let dataObj: any = data;
						if (typeof dataObj === 'string') {
							try { dataObj = JSON.parse(dataObj); } catch(e) { /* use as-is */ }
						}

						if (!dataObj || typeof dataObj !== 'object' || !dataObj.sheets) {
							console.error(LOG, 'invalid spreadsheet data:', dataObj);
							await joplin.views.dialogs.showMessageBox(
								'Spreadsheet data is corrupted or in an unexpected format.\n\n' +
								'Data type: ' + typeof dataObj
							);
							return;
						}
						const dataName = dataObj.name || 'Spreadsheet';
						const sheetCount = (dataObj.sheetOrder && dataObj.sheetOrder.length) || 1;

						// HTML must not contain inline <script> (CSP). Data passes via hidden div;
						// dialogScript.js was registered once via addScript() in onStart.
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
								/* Highlight the cell whose reference is "live" while typing a formula */
								td.ref-target { outline: 2px dashed #1f883d; outline-offset: -2px; background: #dcffe4; }
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
									<input type="hidden" name="state" id="spreadsheet-state" value="">
								</form>
							</div>
							<div class="info">
								💡 Edit cells directly. Use formulas like =SUM(A1:A5), =AVERAGE, =IF. Click "Save &amp; Close" to persist.
							</div>
							<div id="spreadsheet-data">${JSON.stringify(dataObj).replace(/</g, '\\u003c')}</div>
						`;

						await joplin.views.dialogs.setHtml(dialogHandle, html);
						const result = await joplin.views.dialogs.open(dialogHandle);

						if (result.id === 'save') {
							const formData = result.formData;
							let updatedData: any = dataObj;

							// Preferred path: dialogScript.js writes the authoritative in-memory
							// state (including formula cells {f, v}) into a hidden input on every
							// edit. Parse that to preserve formulas exactly as entered.
							const stateStr = formData && formData.spreadsheet && formData.spreadsheet.state;
							let stateParsed: any = null;
							if (stateStr) {
								try { stateParsed = JSON.parse(stateStr); } catch (e) {
									console.warn(LOG, 'failed to parse dialog state, falling back to cell merge:', e);
								}
							}

							if (stateParsed && stateParsed.sheets) {
								updatedData = stateParsed;
							} else if (formData && formData.spreadsheet) {
								// Fallback: merge per-cell form inputs into existing data. Only used
								// if the hidden state input wasn't populated for some reason.
								const firstSheetId = (updatedData.sheetOrder && updatedData.sheetOrder[0]) || Object.keys(updatedData.sheets)[0];
								const existingCellData = updatedData.sheets[firstSheetId].cellData || {};
								const newCellData: any = {};

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

								for (const [key, value] of Object.entries(formData.spreadsheet)) {
									const match = key.match(/^cell-(\d+)-(\d+)$/);
									if (match) {
										const row = parseInt(match[1]);
										const col = parseInt(match[2]);
										if (!newCellData[row]) newCellData[row] = {};
										const existing = newCellData[row][col];
										if (!existing || !existing.f) {
											if (value !== '') {
												newCellData[row][col] = { v: value };
											} else {
												delete newCellData[row][col];
											}
										}
									}
								}
								updatedData.sheets[firstSheetId].cellData = newCellData;
							}

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
						}

					} catch (error) {
						console.error(LOG, 'openUniverSheet failed:', error);
						await joplin.views.dialogs.showMessageBox(`Failed to open editor: ${error.message}`);
					}
				},
			});
		});

		await step('create toolbar button', async () => {
			await joplin.views.toolbarButtons.create(
				'univerSheetButton',
				'insertUniverSheet',
				ToolbarButtonLocation.EditorToolbar
			);
		});

		await step('create menu item', async () => {
			await joplin.views.menuItems.create(
				'univerSheetMenu',
				'insertUniverSheet',
				MenuItemLocation.Tools,
				{ accelerator: 'CmdOrCtrl+Shift+U' }
			);
		});

		console.info(LOG, 'plugin registration complete');
	},
});
