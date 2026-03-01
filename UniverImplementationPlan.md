1. **Univer** is a modern, powerful spreadsheet engine that replaced Luckysheet
2. **Joplin** uses a plugin architecture with TypeScript/JavaScript
3. Univer supports React integration and has a Facade API for easier integration

Let me create an implementation plan and starter code for you:

## Building a Joplin Plugin with Univer Integration

Based on your requirements, here's a comprehensive guide to build a Joplin plugin that embeds Univer spreadsheets in your notes.

### **Project Overview**

You'll create a plugin called "JoplinUniver" that:
- Embeds Univer spreadsheets directly in Joplin notes
- Replaces the deprecated Luckysheet-based JSheets plugin
- Provides Excel-like functionality with formulas and calculations
- Syncs through Dropbox (leveraging Joplin's existing sync)

***

## **Phase 1: Local Project Setup**

### Step 1: Install Prerequisites

First, ensure you have the necessary tools installed:[1]

```bash
# Install Node.js (v16+ recommended)
# Verify installation
node --version
npm --version

# Install Yeoman and Joplin Plugin Generator globally
npm install -g yo generator-joplin
```

### Step 2: Generate Plugin Scaffolding

Create your plugin project:

```bash
# Navigate to your development directory
cd ~/projects  # or wherever you keep projects

# Generate the plugin
yo joplin

# Answer the prompts:
# Plugin name: joplin-univer-sheets
# Plugin description: Embed Univer spreadsheets in Joplin notes
# Author name: Kamlesh Nanda
# Plugin homepage: https://github.com/yourusername/joplin-univer-sheets
```

This creates a project structure:[2]
```
joplin-univer-sheets/
├── src/
│   ├── index.ts          # Plugin entry point
│   └── manifest.json     # Plugin metadata
├── plugin.config.json    # Build configuration
├── package.json
├── tsconfig.json
└── webpack.config.js
```

### Step 3: Install Univer Dependencies

Navigate into your plugin directory and install Univer packages:[3]

```bash
cd joplin-univer-sheets

# Install core Univer packages
npm install @univerjs/core @univerjs/design @univerjs/docs @univerjs/docs-ui @univerjs/engine-formula @univerjs/engine-render @univerjs/sheets @univerjs/sheets-formula @univerjs/sheets-ui @univerjs/ui @univerjs/facade

# Install React (Univer's view layer dependency)
npm install react react-dom

# Install type definitions
npm install -D @types/react @types/react-dom

# Install Univer presets for easier setup
npm install @univerjs/presets
```

***

## **Phase 2: Plugin Implementation**

### Step 4: Create the Univer Integration Module

Create a new file `src/univerSheet.ts`:

```typescript
import { Univer, UniverInstanceType, LocaleType } from '@univerjs/core';
import { defaultTheme } from '@univerjs/design';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import { FUniver } from '@univerjs/facade';

export class UniverSheetManager {
    private univer: Univer | null = null;
    private univerAPI: FUniver | null = null;

    /**
     * Initialize a Univer instance in the given container
     */
    async initialize(container: HTMLElement, initialData?: any): Promise<FUniver> {
        // Create Univer instance
        this.univer = new Univer({
            theme: defaultTheme,
            locale: LocaleType.EN_US,
            locales: {
                // You can add more locales here
            },
        });

        // Register plugins
        this.univer.registerPlugin(UniverRenderEnginePlugin);
        this.univer.registerPlugin(UniverFormulaEnginePlugin);
        this.univer.registerPlugin(UniverUIPlugin, {
            container: container,
            header: true,
            toolbar: true,
            footer: true,
        });

        // Register sheet plugins
        this.univer.registerPlugin(UniverDocsPlugin);
        this.univer.registerPlugin(UniverDocsUIPlugin);
        this.univer.registerPlugin(UniverSheetsPlugin);
        this.univer.registerPlugin(UniverSheetsUIPlugin);
        this.univer.registerPlugin(UniverSheetsFormulaPlugin);

        // Create workbook
        this.univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialData || this.getDefaultWorkbookData());

        // Get Facade API
        this.univerAPI = FUniver.newAPI(this.univer);

        return this.univerAPI;
    }

    /**
     * Get default workbook structure
     */
    private getDefaultWorkbookData() {
        return {
            id: 'workbook-01',
            name: 'New Spreadsheet',
            sheetOrder: ['sheet-01'],
            appVersion: '0.1.0',
            locale: LocaleType.EN_US,
            styles: {},
            sheets: {
                'sheet-01': {
                    id: 'sheet-01',
                    name: 'Sheet1',
                    cellData: {
                        0: {
                            0: { v: 'Hello' },
                            1: { v: 'Univer' },
                        },
                    },
                    rowCount: 100,
                    columnCount: 26,
                    defaultColumnWidth: 93,
                    defaultRowHeight: 27,
                },
            },
        };
    }

    /**
     * Get current workbook data for saving
     */
    getWorkbookData(): any {
        if (!this.univerAPI) return null;
        
        const activeWorkbook = this.univerAPI.getActiveWorkbook();
        return activeWorkbook?.save();
    }

    /**
     * Dispose the Univer instance
     */
    dispose() {
        if (this.univer) {
            this.univer.dispose();
            this.univer = null;
            this.univerAPI = null;
        }
    }
}
```

### Step 5: Create Content Script for Note Rendering

Create `src/contentScript.ts`:

```typescript
import { UniverSheetManager } from './univerSheet';

// This runs in the note viewer context
export default function(context: any) {
    return {
        plugin: function(markdownIt: any, _options: any) {
            const defaultRender = markdownIt.renderer.rules.fence || function(tokens: any, idx: any, options: any, env: any, self: any) {
                return self.renderToken(tokens, idx, options);
            };

            markdownIt.renderer.rules.fence = function(tokens: any, idx: any, options: any, env: any, self: any) {
                const token = tokens[idx];
                
                // Check if this is a univer code block
                if (token.info === 'univer' || token.info === 'univer-sheet') {
                    try {
                        const data = JSON.parse(token.content);
                        const id = `univer-${Math.random().toString(36).substr(2, 9)}`;
                        
                        // Create container
                        const html = `
                            <div class="univer-container" id="${id}" data-univer-data='${JSON.stringify(data)}' style="width: 100%; height: 600px; border: 1px solid #ccc; border-radius: 4px;">
                                <div class="univer-loading">Loading spreadsheet...</div>
                            </div>
                        `;
                        
                        // Schedule initialization after DOM is ready
                        setTimeout(() => initializeUniver(id), 100);
                        
                        return html;
                    } catch (e) {
                        console.error('Failed to parse Univer data:', e);
                        return `<pre>Error loading Univer spreadsheet: ${e.message}</pre>`;
                    }
                }
                
                return defaultRender(tokens, idx, options, env, self);
            };
        },
        assets: function() {
            return [
                { name: './univer-bundle.css' }
            ];
        },
    };
}

// Initialize Univer instance
async function initializeUniver(containerId: string) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    const dataAttr = container.getAttribute('data-univer-data');
    if (!dataAttr) return;
    
    try {
        const data = JSON.parse(dataAttr);
        const manager = new UniverSheetManager();
        await manager.initialize(container, data);
        
        // Remove loading message
        const loading = container.querySelector('.univer-loading');
        if (loading) loading.remove();
        
        // Store manager for later cleanup
        (window as any)[`univerManager_${containerId}`] = manager;
    } catch (e) {
        console.error('Failed to initialize Univer:', e);
        container.innerHTML = `<div style="color: red;">Failed to initialize spreadsheet: ${e.message}</div>`;
    }
}
```

### Step 6: Update Main Plugin File

Edit `src/index.ts`:

```typescript
import joplin from 'api';
import { MenuItemLocation, ToolbarButtonLocation } from 'api/types';

joplin.plugins.register({
    onStart: async function() {
        console.log('Joplin Univer Sheets plugin started');

        // Register content script for rendering Univer blocks in notes
        await joplin.contentScripts.register(
            'univerSheet',
            'univerSheet',
            './contentScript.js'
        );

        // Register command to insert Univer spreadsheet
        await joplin.commands.register({
            name: 'insertUniverSheet',
            label: 'Insert Univer Spreadsheet',
            execute: async () => {
                const noteId = await joplin.workspace.selectedNote();
                if (!noteId) return;

                // Default empty spreadsheet data
                const defaultData = {
                    id: `sheet-${Date.now()}`,
                    name: 'Spreadsheet',
                    sheetOrder: ['sheet-01'],
                    sheets: {
                        'sheet-01': {
                            id: 'sheet-01',
                            name: 'Sheet1',
                            cellData: {},
                            rowCount: 100,
                            columnCount: 26,
                        },
                    },
                };

                const markdown = `\n\`\`\`univer-sheet\n${JSON.stringify(defaultData, null, 2)}\n\`\`\`\n`;
                await joplin.commands.execute('insertText', markdown);
            },
        });

        // Add toolbar button
        await joplin.views.toolbarButtons.create(
            'univerSheetButton',
            'insertUniverSheet',
            ToolbarButtonLocation.EditorToolbar
        );

        // Add menu item
        await joplin.views.menuItems.create(
            'univerSheetMenu',
            'insertUniverSheet',
            MenuItemLocation.Tools,
            { accelerator: 'CmdOrCtrl+Shift+U' }
        );

        console.log('Univer Sheets plugin registration complete');
    },
});
```

### Step 7: Configure Build System

Update `plugin.config.json` to include content scripts:

```json
{
    "extraScripts": [
        {
            "path": "./src/contentScript.ts",
            "outputPath": "contentScript.js",
            "assetType": "contentScript"
        }
    ]
}
```

Update `package.json` to include Univer styles:

```json
{
  "scripts": {
    "dist": "webpack --config webpack.config.js",
    "prepare": "npm run dist",
    "copy-styles": "cp node_modules/@univerjs/design/lib/index.css dist/univer-bundle.css"
  }
}
```

***

## **Phase 3: Build and Test**

### Step 8: Build the Plugin

```bash
# Install all dependencies
npm install

# Build the plugin
npm run dist

# This creates:
# - dist/ folder with compiled plugin
# - *.jpl file (Joplin Plugin) at the root
```

### Step 9: Install in Joplin

**Option A: Development Mode (Recommended for testing)**[1]

1. Open Joplin in development mode:
   ```bash
   # On macOS/Linux
   /Applications/Joplin.app/Contents/MacOS/Joplin --profile ~/.config/joplin-dev
   
   # On Windows
   Joplin.exe --profile %USERPROFILE%\.config\joplin-dev
   ```

2. In Joplin, go to **Tools > Options > Plugins**

3. Under "Advanced Settings", add your plugin path in **Development plugins**:
   ```
   /path/to/joplin-univer-sheets
   ```

4. Restart Joplin development instance

**Option B: Install Built Plugin**

1. Build the `.jpl` file: `npm run dist`
2. In Joplin: **Tools > Options > Plugins**
3. Click "Install plugin" and select the `.jpl` file
4. Restart Joplin

### Step 10: Test the Plugin

1. Create a new note in Joplin

2. Click the Univer toolbar button (or press `Ctrl+Shift+U` / `Cmd+Shift+U`)

3. This inserts a code block like:
   ````markdown
   ```univer-sheet
   {
     "id": "sheet-1234567890",
     "name": "Spreadsheet",
     ...
   }
   ```
   ````

4. Switch to view mode - you should see the interactive Univer spreadsheet

5. Test spreadsheet features:
   - Enter data in cells
   - Use formulas (=SUM(A1:A10))
   - Format cells
   - Add/remove sheets

***

## **Phase 4: Advanced Features**

### Feature 1: Auto-Save to Note Content

Add to `src/index.ts`:

```typescript
// Register periodic save
let saveInterval: NodeJS.Timer;

await joplin.workspace.onNoteSelectionChange(async () => {
    if (saveInterval) clearInterval(saveInterval);
    
    // Auto-save every 30 seconds
    saveInterval = setInterval(async () => {
        const note = await joplin.workspace.selectedNote();
        if (note) {
            // Trigger save event in content script
            await joplin.commands.execute('editor.execCommand', {
                name: 'univerAutoSave'
            });
        }
    }, 30000);
});
```

### Feature 2: Import/Export Excel Files

Add import command:

```typescript
await joplin.commands.register({
    name: 'importExcelToUniver',
    label: 'Import Excel File',
    execute: async () => {
        const filePath = await joplin.require('fs').showOpenDialog({
            filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }]
        });
        
        if (filePath) {
            // Read file and convert to Univer format
            // Use @univerjs/sheets-plugin-import
            const data = await convertExcelToUniver(filePath);
            const markdown = `\n\`\`\`univer-sheet\n${JSON.stringify(data, null, 2)}\n\`\`\`\n`;
            await joplin.commands.execute('insertText', markdown);
        }
    },
});
```

### Feature 3: Collaborative Editing Indicator

While Joplin syncs through Dropbox, you can add indicators for recently modified sheets:

```typescript
// Add metadata to track modifications
const metadata = {
    lastModified: new Date().toISOString(),
    modifiedBy: await joplin.settings.value('sync.username') || 'unknown',
};
```

***

## **Integration with Joplin Sync**

The beauty of this approach is that **Joplin's existing Dropbox sync handles everything**:[4]

1. **Spreadsheet data is stored as JSON in the markdown code block**
2. **Joplin syncs the entire note (including the JSON) through Dropbox**
3. **On other devices, the content script renders the spreadsheet from the JSON**
4. **No additional sync mechanism needed!**

The workflow:
```
Desktop → Edit spreadsheet → JSON updated in note → Joplin syncs note → 
Dropbox sync → Other devices pull changes → Spreadsheet renders with new data
```

***

## **Troubleshooting Common Issues**

### Issue 1: Univer CSS not loading

**Solution**: Ensure you're copying Univer styles:

```bash
npm run copy-styles
```

Add to `contentScript.ts` assets:

```typescript
assets: function() {
    return [
        { name: './univer-bundle.css' }
    ];
}
```

### Issue 2: React version conflicts

Univer requires React 18.3.1. If you see errors:[5]

```bash
npm install react@18.3.1 react-dom@18.3.1
```

### Issue 3: Plugin not loading

Check Joplin's log:
- **Help > Toggle Development Tools > Console**
- Look for plugin errors

### Issue 4: Large spreadsheets causing performance issues

Limit initial render size:

```typescript
sheets: {
    'sheet-01': {
        rowCount: 50,  // Instead of 1000
        columnCount: 26,
    },
}
```

***

## **Next Steps & Enhancements**

Once the basic plugin works, consider:

1. **Formula Library**: Add custom formulas specific to your workflow
2. **Templates**: Pre-built spreadsheet templates for budgets, tracking, etc.
3. **Chart Support**: Enable Univer's charting capabilities[6]
4. **Cell Formatting**: Rich formatting options
5. **Data Validation**: Dropdown lists, data validation rules
6. **Keyboard Shortcuts**: Excel-like navigation shortcuts
7. **Mobile Support**: Test and optimize for Joplin mobile apps
8. **Export Options**: Export to PDF, CSV, Excel

***

## **Resources & References**

- **Univer Documentation**: https://docs.univer.ai[7]
- **Univer GitHub**: https://github.com/dream-num/univer[6]
- **Joplin Plugin API**: https://joplinapp.org/help/api/get_started/plugins/[1]
- **Univer React Integration**: https://docs.univer.ai/guides/docs/getting-started/integrations/react[5]
- **Univer React Demo**: https://github.com/awesome-univer/sheets-react-demo[8]

***

Would you like me to:

1. **Create the complete project files** as downloadable artifacts?
2. **Help you debug specific integration issues** once you start building?
3. **Create an Excel import/export module** for seamless file handling?
4. **Build a custom toolbar** with finance-specific formulas for your use case?

Let me know which direction you'd like to go, and I'll provide more detailed assistance!

Sources
[1] Getting started with plugin development - Joplin https://joplinapp.org/help/api/get_started/plugins/
[2] joplin-plugin-note-overview/GENERATOR_DOC.md at master - GitHub https://github.com/JackGruber/joplin-plugin-note-overview/blob/master/GENERATOR_DOC.md
[3] Installation & Basic Usage - Univer https://docs.univer.ai/guides/sheets/getting-started/installation
[4] Joplin Development | James' Digital Garden https://wiki.jamesravey.me/books/node-and-typescript/page/joplin-development
[5] React - Univer https://docs.univer.ai/guides/docs/getting-started/integrations/react
[6] dream-num/univer: Build AI-native spreadsheets. Univer is ... - GitHub https://github.com/dream-num/univer
[7] Univer https://docs.univer.ai
[8] The demo of using Univer sheets with React Hooks - GitHub https://github.com/awesome-univer/sheets-react-demo
[9] Former Feishu Spreadsheet Tech Lead's Ventures https://eu.36kr.com/en/p/3675553046831752
[10] Integrations - Univer https://docs.univer.ai/guides/sheets/getting-started/integrations
[11] How To Embed A Spreadsheet To Jotform (Tutorial 2026) - YouTube https://www.youtube.com/watch?v=_I2sgOhPMkk
[12] awesome-univer/sheets-react-facade-demo - GitHub https://github.com/awesome-univer/sheets-react-facade-demo
[13] Tutorial for getting started on Joplin plugin ecosystem? - Development https://discourse.joplinapp.org/t/tutorial-for-getting-started-on-joplin-plugin-ecosystem/11290
[14] Open-Source Spreadsheets: The Golden Gateway Between AI and ... https://www.reddit.com/r/opensource/comments/1kx8whq/opensource_spreadsheets_the_golden_gateway/
[15] Plugin : Template - Joplin - YouTube https://www.youtube.com/watch?v=pw9DVitXjfQ
