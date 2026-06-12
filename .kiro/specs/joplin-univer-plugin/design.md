# Design Document: Joplin Univer Plugin

## 1. Overview

This document describes the technical design for a Joplin plugin that integrates Univer spreadsheet functionality. The plugin enables users to embed interactive spreadsheets within Joplin notes, providing Excel-like capabilities including formulas, cell formatting, and data persistence through Joplin's existing sync mechanisms.

## 2. Architecture

### 2.1 High-Level Architecture

````
┌─────────────────────────────────────────────────────────┐
│                    Joplin Application                    │
│  ┌───────────────────────────────────────────────────┐  │
│  │              Plugin Main Process                   │  │
│  │  - Command Registration                           │  │
│  │  - Toolbar/Menu Integration                       │  │
│  │  - Content Script Registration                    │  │
│  └───────────────────────────────────────────────────┘  │
│                          │                               │
│  ┌───────────────────────────────────────────────────┐  │
│  │           Content Script (Note Renderer)          │  │
│  │  - Markdown Fence Block Parser                    │  │
│  │  - Univer Instance Manager                        │  │
│  │  - DOM Manipulation                               │  │
│  └───────────────────────────────────────────────────┘  │
│                          │                               │
│  ┌───────────────────────────────────────────────────┐  │
│  │              Univer SDK Layer                     │  │
│  │  - Core Engine                                    │  │
│  │  - Sheets Plugin                                  │  │
│  │  - Formula Engine                                 │  │
│  │  - UI Components                                  │  │
│  └───────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│              Joplin Note Storage (Markdown)              │
│  ```univer-sheet                                         │
│  { "id": "...", "sheets": {...}, "cellData": {...} }    │
│  ```                                                     │
└─────────────────────────────────────────────────────────┘
````

### 2.2 Component Breakdown

#### 2.2.1 Plugin Main Module (`src/index.ts`)

- Registers plugin with Joplin API
- Creates toolbar button, menu items, and keyboard shortcuts
- Registers content script for note rendering
- Handles "Insert Spreadsheet" command

#### 2.2.2 Content Script (`src/contentScript.ts`)

- Intercepts Markdown rendering
- Detects `univer-sheet` fence blocks
- Initializes Univer instances in the DOM
- Manages lifecycle of spreadsheet components

#### 2.2.3 Univer Manager (`src/univerSheet.ts`)

- Encapsulates Univer SDK initialization
- Manages workbook data structure
- Provides serialization/deserialization methods
- Handles plugin registration and configuration

## 3. Data Model

### 3.1 Spreadsheet Storage Format

Spreadsheets are stored as JSON within Markdown fence blocks:

````markdown
```univer-sheet
{
  "id": "workbook-1234567890",
  "name": "Budget 2024",
  "sheetOrder": ["sheet-01"],
  "appVersion": "0.1.0",
  "locale": "en-US",
  "styles": {},
  "sheets": {
    "sheet-01": {
      "id": "sheet-01",
      "name": "Sheet1",
      "cellData": {
        "0": {
          "0": { "v": "Item", "s": "style-id" },
          "1": { "v": "Cost" }
        },
        "1": {
          "0": { "v": "Coffee" },
          "1": { "v": 5.99 }
        }
      },
      "rowCount": 100,
      "columnCount": 26,
      "defaultColumnWidth": 93,
      "defaultRowHeight": 27
    }
  }
}
```
````

````

### 3.2 Cell Data Structure

Each cell contains:
- `v`: Value (string, number, or formula)
- `s`: Style ID (optional)
- `f`: Formula string (optional)
- `t`: Cell type (optional)

### 3.3 Default Workbook Template

New spreadsheets are initialized with:
- Single sheet named "Sheet1"
- 100 rows × 26 columns
- Default cell dimensions
- Empty cell data

## 4. Key Technical Decisions

### 4.1 Storage Strategy

**Decision**: Store spreadsheet data as JSON within Markdown fence blocks

**Rationale**:
- Leverages Joplin's existing sync infrastructure
- No additional backend or database required
- Text-based format compatible with all sync providers (Dropbox, Joplin Cloud, etc.)
- Version control friendly
- Human-readable for debugging

**Trade-offs**:
- Large spreadsheets increase note size
- No binary compression
- Potential merge conflicts with concurrent edits

### 4.2 Rendering Approach

**Decision**: Use content scripts with Markdown-it fence block interception

**Rationale**:
- Integrates seamlessly with Joplin's rendering pipeline
- Spreadsheets render inline at correct position
- Supports multiple spreadsheets per note
- No modification to Joplin core required

### 4.3 Univer SDK Integration

**Decision**: Use Univer Facade API with modular plugin architecture

**Rationale**:
- Facade API simplifies integration
- Modular plugins allow selective feature loading
- Open-source and actively maintained
- Modern replacement for deprecated Luckysheet
- Supports formulas, formatting, and Excel-like features

### 4.4 Build System

**Decision**: Use Webpack with TypeScript and ts-loader

**Rationale**:
- Standard Joplin plugin build system
- TypeScript provides type safety
- Webpack bundles all dependencies into single .jpl file
- Supports content script compilation

## 5. Implementation Details

### 5.1 Plugin Initialization Flow

1. Joplin loads plugin from .jpl file
2. Plugin registers with Joplin API
3. Content script registered for note rendering
4. Commands, toolbar buttons, and menu items created
5. Plugin enters ready state

### 5.2 Spreadsheet Creation Flow

1. User clicks toolbar button or uses keyboard shortcut
2. Plugin generates default workbook JSON
3. JSON wrapped in ```univer-sheet fence block
4. Fence block inserted at cursor position in note
5. Note re-renders, triggering content script

### 5.3 Spreadsheet Rendering Flow

1. Markdown-it parser encounters `univer-sheet` fence block
2. Content script intercepts fence block rendering
3. JSON data extracted and parsed
4. Container div created with unique ID
5. Univer instance initialized in container
6. Workbook data loaded into Univer
7. Interactive spreadsheet displayed

### 5.4 Data Persistence Flow

1. User edits spreadsheet cells
2. Changes stored in Univer's internal state
3. On blur or periodic interval, data serialized
4. JSON in fence block updated
5. Joplin saves note with updated content
6. Sync mechanism propagates changes

### 5.5 Error Handling Strategy

- **SDK Load Failure**: Display error message, prevent initialization
- **Parse Error**: Show warning, display raw JSON
- **Formula Error**: Display error indicator in cell
- **Sync Conflict**: Joplin's conflict resolution applies
- **Memory Issues**: Limit default spreadsheet size

## 6. Dependencies

### 6.1 Core Dependencies

```json
{
  "@univerjs/core": "^0.5.0",
  "@univerjs/design": "^0.5.0",
  "@univerjs/engine-formula": "^0.5.0",
  "@univerjs/engine-render": "^0.5.0",
  "@univerjs/sheets": "^0.5.0",
  "@univerjs/sheets-formula": "^0.5.0",
  "@univerjs/sheets-ui": "^0.5.0",
  "@univerjs/ui": "^0.5.0",
  "@univerjs/facade": "^0.5.0",
  "@univerjs/docs": "^0.5.0",
  "@univerjs/docs-ui": "^0.5.0",
  "react": "^18.3.1",
  "react-dom": "^18.3.1"
}
````

### 6.2 Development Dependencies

```json
{
    "typescript": "^5.0.0",
    "webpack": "^5.0.0",
    "webpack-cli": "^5.0.0",
    "ts-loader": "^9.0.0",
    "@types/react": "^18.0.0",
    "@types/react-dom": "^18.0.0"
}
```

## 7. API Interfaces

### 7.1 UniverSheetManager Interface

```typescript
class UniverSheetManager {
    // Initialize Univer instance in container
    async initialize(container: HTMLElement, initialData?: WorkbookData): Promise<FUniver>;

    // Get current workbook data for saving
    getWorkbookData(): WorkbookData | null;

    // Clean up and dispose instance
    dispose(): void;

    // Get default workbook structure
    private getDefaultWorkbookData(): WorkbookData;
}
```

### 7.2 Content Script Interface

```typescript
export default function(context: any) {
  return {
    // Markdown-it plugin
    plugin: function(markdownIt: any, options: any)

    // CSS assets to load
    assets: function(): Array<{name: string}>
  }
}
```

### 7.3 Joplin Plugin Interface

```typescript
joplin.plugins.register({
  onStart: async function() {
    // Register content script
    await joplin.contentScripts.register(...)

    // Register commands
    await joplin.commands.register(...)

    // Create UI elements
    await joplin.views.toolbarButtons.create(...)
    await joplin.views.menuItems.create(...)
  }
})
```

## 8. Security Considerations

### 8.1 Input Validation

- Validate JSON structure before parsing
- Sanitize cell values to prevent XSS
- Limit spreadsheet dimensions to prevent DoS

### 8.2 Resource Limits

- Maximum spreadsheet size: 1000 rows × 100 columns
- Maximum cell content length: 10KB
- Maximum workbook JSON size: 5MB

### 8.3 Formula Security

- Formulas execute in sandboxed environment
- No access to file system or network
- No eval() or code execution

## 9. Performance Considerations

### 9.1 Rendering Optimization

- Lazy initialization of Univer instances
- Virtual scrolling for large spreadsheets
- Debounced save operations

### 9.2 Memory Management

- Dispose Univer instances when notes close
- Limit concurrent spreadsheet instances
- Clear unused workbook data

### 9.3 Sync Optimization

- Compress JSON where possible
- Delta updates for large spreadsheets (future)
- Throttle save operations

## 10. Testing Strategy

### 10.1 Unit Tests

- UniverSheetManager initialization
- Data serialization/deserialization
- Default workbook generation

### 10.2 Integration Tests

- Plugin registration with Joplin
- Content script rendering
- Command execution

### 10.3 Manual Testing

- Create spreadsheet in note
- Edit cells and formulas
- Save and reload note
- Sync across devices
- Multiple spreadsheets per note

## 11. Future Enhancements

### 11.1 Phase 2 Features

- Auto-save with configurable interval
- Import/export Excel files
- Custom cell formatting
- Chart support

### 11.2 Phase 3 Features

- Collaborative editing indicators
- Spreadsheet templates
- Advanced formula library
- Mobile optimization

### 11.3 Phase 4 Features

- Real-time collaboration
- Cloud-based calculation
- Plugin API for extensions
- Custom functions

## 12. Deployment

### 12.1 Build Process

```bash
npm install          # Install dependencies
npm run dist         # Build plugin
```

Output: `publish/com.kamleshnanda.joplin-univer-plugin.jpl`

### 12.2 Installation Methods

**Development Mode**:

- Add plugin directory to Joplin development plugins
- Restart Joplin

**Production Mode**:

- Install .jpl file via Joplin plugin manager
- Restart Joplin

### 12.3 Distribution

- Publish to npm with `joplin-plugin-` prefix
- Include keywords: `joplin-plugin`, `Excel`, `Spreadsheet`, `Univer`, `Kamlesh`
- Automatic inclusion in Joplin plugin repository

## 13. Maintenance

### 13.1 Version Management

- Follow semantic versioning
- Update manifest.json and package.json in sync
- Use `npm run updateVersion` script

### 13.2 Dependency Updates

- Monitor Univer SDK releases
- Test compatibility before upgrading
- Document breaking changes

### 13.3 Issue Tracking

- GitHub issues for bug reports
- Feature requests via discussions
- Security issues via private disclosure

## 14. License

MIT License - Compatible with Joplin and Univer open-source licenses
