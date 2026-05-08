/**
 * Dialog script for rendering and editing spreadsheet data.
 * Loaded via joplin.views.dialogs.addScript() — runs inside the dialog webview.
 * 
 * Data is passed from the plugin via a hidden <div id="spreadsheet-data">
 * containing JSON-encoded spreadsheet data.
 * 
 * VERSION: 2.4.0 — adds formula evaluation engine + row/column manipulation
 * 
 * Supported formulas:
 *   =A1+B1, =A1*2, =A1/B1, =A1-B1
 *   =SUM(A1:C1), =AVERAGE(A1:A10), =COUNT(A1:A10)
 *   =MIN(A1:A10), =MAX(A1:A10)
 *   =IF(A1>0,"yes","no")
 *   =CONCAT(A1,B1), =ABS(A1), =ROUND(A1,2)
 *   =UPPER(A1), =LOWER(A1), =LEN(A1), =TRIM(A1)
 */
(function() {
    // Limits to prevent oversized data
    var MAX_CELL_LENGTH = 10000;
    var MAX_ROWS = 500;
    var MAX_COLS = 52;

    var _firstSheetId = null;

    // ── Cell mode state (Excel-like) ──
    // _editMode: false = "selected" (whole content highlighted, typing replaces).
    //            true  = "edit"     (cursor in text, arrows behave per content type).
    var _editMode = false;

    // ── Reference-picker state ──
    // While entering a formula (value starts with "="), arrow keys and clicks insert
    // a cell reference at the cursor. After insertion, the reference is "live": the
    // next arrow key REPLACES it (matching Excel). Any non-arrow keystroke commits
    // the live reference (so the next arrow inserts a new ref at the cursor).
    var _refLive = false;
    var _refRow = -1;        // current target row (live ref points here)
    var _refCol = -1;        // current target col
    var _refStart = -1;      // slice [start,end) in active input where the ref text sits
    var _refEnd = -1;

    // Button styles (inline since CSP blocks <style> in dynamic HTML)
    var _btnStyle = 'padding:4px 10px;font-size:12px;border:1px solid #d0d7de;border-radius:4px;background:#f6f8fa;color:#24292f;cursor:pointer;';
    var _btnStyleDanger = 'padding:4px 10px;font-size:12px;border:1px solid #d0d7de;border-radius:4px;background:#fff5f5;color:#cf222e;cursor:pointer;';

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    function init() {
        var dataEl = document.getElementById('spreadsheet-data');
        if (!dataEl) { showError('No data element found'); return; }

        var raw = dataEl.textContent || dataEl.innerText || '';
        var data = null;
        try { data = JSON.parse(raw); } catch(e) { showError('Parse error: ' + e.message); return; }
        if (typeof data === 'string') { try { data = JSON.parse(data); } catch(e) {} }
        if (!data || typeof data !== 'object') { showError('Invalid data type: ' + typeof data); return; }

        window.spreadsheetData = data;
        renderSpreadsheet(data);
        syncState(); // initial state for the hidden input
        // Excel-like default: cell A1 is selected on open so keyboard works immediately.
        focusCellAt(0, 0, /* selectMode */ true);
    }

    // Mirror window.spreadsheetData into the hidden #spreadsheet-state input so the
    // plugin process receives authoritative cell data (including formulas as {f, v})
    // on save, instead of having to reconstruct from per-cell input values.
    function syncState() {
        var stateEl = document.getElementById('spreadsheet-state');
        if (!stateEl || !window.spreadsheetData) return;
        try {
            stateEl.value = JSON.stringify(window.spreadsheetData);
        } catch(e) {
            console.error('syncState failed:', e);
        }
    }

    function showError(msg) {
        var el = document.getElementById('spreadsheet-editor');
        if (el) el.innerHTML = '<p style="color:red;padding:12px;">' + escapeHtml(msg) + '</p>';
    }

    function showWarning(msg) {
        var container = document.querySelector('.spreadsheet-container');
        if (!container) return;
        var existing = document.getElementById('spreadsheet-warning');
        if (existing) existing.remove();
        var warn = document.createElement('div');
        warn.id = 'spreadsheet-warning';
        warn.style.cssText = 'background:#fff3cd;color:#856404;padding:8px 12px;border-radius:4px;margin-bottom:8px;font-size:13px;';
        warn.textContent = '\u26a0\ufe0f ' + msg;
        container.insertBefore(warn, container.firstChild);
        setTimeout(function() { if (warn.parentNode) warn.remove(); }, 5000);
    }

    // ── Column helpers ──
    function colToIndex(col) {
        col = col.toUpperCase();
        var idx = 0;
        for (var i = 0; i < col.length; i++) {
            idx = idx * 26 + (col.charCodeAt(i) - 64);
        }
        return idx - 1;
    }

    function indexToCol(idx) {
        var s = '';
        idx++;
        while (idx > 0) { idx--; s = String.fromCharCode(65 + (idx % 26)) + s; idx = Math.floor(idx / 26); }
        return s;
    }

    // ── Formula engine ──
    // Resolves a cell reference like "A1" to its raw value from cellData
    function getCellRaw(cellData, ref) {
        var m = ref.toUpperCase().match(/^([A-Z]+)(\d+)$/);
        if (!m) return undefined;
        var col = colToIndex(m[1]);
        var row = parseInt(m[2]) - 1;
        var rd = cellData[row] || cellData[String(row)];
        if (!rd) return undefined;
        var cell = rd[col] || rd[String(col)];
        if (!cell) return undefined;
        return cell.f !== undefined ? cell.f : (cell.v !== undefined ? cell.v : undefined);
    }

    // Get the computed numeric value of a cell (evaluates formulas recursively)
    function getCellValue(cellData, ref, visited) {
        if (!visited) visited = {};
        var key = ref.toUpperCase();
        if (visited[key]) return '#CIRC!';
        visited[key] = true;
        var raw = getCellRaw(cellData, ref);
        if (raw === undefined || raw === '') return 0;
        if (typeof raw === 'string' && raw.charAt(0) === '=') {
            return evaluateFormula(raw, cellData, visited);
        }
        var n = Number(raw);
        return isNaN(n) ? raw : n;
    }

    // Expand a range like "A1:C3" into an array of cell references
    function expandRange(rangeStr) {
        var parts = rangeStr.toUpperCase().split(':');
        if (parts.length !== 2) return [rangeStr];
        var m1 = parts[0].match(/^([A-Z]+)(\d+)$/);
        var m2 = parts[1].match(/^([A-Z]+)(\d+)$/);
        if (!m1 || !m2) return [rangeStr];
        var c1 = colToIndex(m1[1]), r1 = parseInt(m1[2]);
        var c2 = colToIndex(m2[1]), r2 = parseInt(m2[2]);
        var refs = [];
        for (var r = Math.min(r1,r2); r <= Math.max(r1,r2); r++) {
            for (var c = Math.min(c1,c2); c <= Math.max(c1,c2); c++) {
                refs.push(indexToCol(c) + r);
            }
        }
        return refs;
    }

    // Tokenize formula arguments, respecting nested parens and strings
    function splitArgs(str) {
        var args = [], depth = 0, current = '', inStr = false, strChar = '';
        for (var i = 0; i < str.length; i++) {
            var ch = str[i];
            if (inStr) {
                current += ch;
                if (ch === strChar) inStr = false;
            } else if (ch === '"' || ch === "'") {
                inStr = true; strChar = ch; current += ch;
            } else if (ch === '(') {
                depth++; current += ch;
            } else if (ch === ')') {
                depth--; current += ch;
            } else if (ch === ',' && depth === 0) {
                args.push(current.trim()); current = '';
            } else {
                current += ch;
            }
        }
        if (current.trim()) args.push(current.trim());
        return args;
    }

    // Collect numeric values from a list of refs/ranges
    function collectValues(args, cellData, visited) {
        var vals = [];
        for (var i = 0; i < args.length; i++) {
            var arg = args[i].trim();
            if (arg.indexOf(':') !== -1) {
                var refs = expandRange(arg);
                for (var j = 0; j < refs.length; j++) {
                    var v = getCellValue(cellData, refs[j], Object.assign({}, visited));
                    if (typeof v === 'number') vals.push(v);
                }
            } else if (arg.match(/^[A-Z]+\d+$/i)) {
                var v2 = getCellValue(cellData, arg, Object.assign({}, visited));
                if (typeof v2 === 'number') vals.push(v2);
            } else {
                var n = Number(arg);
                if (!isNaN(n)) vals.push(n);
            }
        }
        return vals;
    }

    // Evaluate a single expression token (number, string literal, cell ref, or sub-formula)
    function evalToken(token, cellData, visited) {
        token = token.trim();
        if (token === '') return 0;
        // String literal
        if ((token.charAt(0) === '"' && token.charAt(token.length-1) === '"') ||
            (token.charAt(0) === "'" && token.charAt(token.length-1) === "'")) {
            return token.substring(1, token.length - 1);
        }
        // Function call
        var fnMatch = token.match(/^([A-Z]+)\((.+)\)$/i);
        if (fnMatch) {
            return evalFunction(fnMatch[1].toUpperCase(), fnMatch[2], cellData, visited);
        }
        // Cell reference
        if (token.match(/^[A-Z]+\d+$/i)) {
            return getCellValue(cellData, token, Object.assign({}, visited));
        }
        // Number
        var n = Number(token);
        if (!isNaN(n)) return n;
        return token;
    }

    // Evaluate built-in functions
    function evalFunction(name, argsStr, cellData, visited) {
        var args = splitArgs(argsStr);
        switch (name) {
            case 'SUM': {
                var vals = collectValues(args, cellData, visited);
                var s = 0; for (var i = 0; i < vals.length; i++) s += vals[i];
                return s;
            }
            case 'AVERAGE': {
                var vals = collectValues(args, cellData, visited);
                if (vals.length === 0) return '#DIV/0!';
                var s = 0; for (var i = 0; i < vals.length; i++) s += vals[i];
                return s / vals.length;
            }
            case 'COUNT': {
                return collectValues(args, cellData, visited).length;
            }
            case 'MIN': {
                var vals = collectValues(args, cellData, visited);
                if (vals.length === 0) return 0;
                return Math.min.apply(null, vals);
            }
            case 'MAX': {
                var vals = collectValues(args, cellData, visited);
                if (vals.length === 0) return 0;
                return Math.max.apply(null, vals);
            }
            case 'ABS': {
                var v = evalToken(args[0], cellData, visited);
                return typeof v === 'number' ? Math.abs(v) : '#VALUE!';
            }
            case 'ROUND': {
                var v = evalToken(args[0], cellData, visited);
                var d = args[1] ? Number(evalToken(args[1], cellData, visited)) : 0;
                if (typeof v !== 'number') return '#VALUE!';
                var f = Math.pow(10, d);
                return Math.round(v * f) / f;
            }
            case 'IF': {
                if (args.length < 2) return '#VALUE!';
                var cond = evalExpression(args[0], cellData, visited);
                if (cond && cond !== 0 && cond !== '' && cond !== false) {
                    return evalToken(args[1], cellData, visited);
                }
                return args[2] ? evalToken(args[2], cellData, visited) : '';
            }
            case 'CONCAT': {
                var result = '';
                for (var i = 0; i < args.length; i++) {
                    result += String(evalToken(args[i], cellData, visited));
                }
                return result;
            }
            case 'UPPER': {
                return String(evalToken(args[0], cellData, visited)).toUpperCase();
            }
            case 'LOWER': {
                return String(evalToken(args[0], cellData, visited)).toLowerCase();
            }
            case 'LEN': {
                return String(evalToken(args[0], cellData, visited)).length;
            }
            case 'TRIM': {
                return String(evalToken(args[0], cellData, visited)).trim();
            }
            default:
                return '#NAME?';
        }
    }

    // Simple expression evaluator: handles +, -, *, /, comparisons, parens, function calls
    function evalExpression(expr, cellData, visited) {
        expr = expr.trim();
        // Try to parse as a simple infix expression
        // Tokenize into numbers, cell refs, operators, parens, function calls, strings
        var tokens = tokenize(expr);
        if (tokens.length === 0) return 0;
        if (tokens.length === 1) return evalToken(tokens[0], cellData, visited);
        return evalTokenList(tokens, cellData, visited);
    }

    function tokenize(expr) {
        var tokens = [], i = 0, len = expr.length;
        while (i < len) {
            var ch = expr[i];
            if (ch === ' ') { i++; continue; }
            // String literal
            if (ch === '"' || ch === "'") {
                var q = ch, s = ch; i++;
                while (i < len && expr[i] !== q) { s += expr[i]; i++; }
                if (i < len) { s += expr[i]; i++; }
                tokens.push(s);
            }
            // Number
            else if ((ch >= '0' && ch <= '9') || (ch === '.' && i+1 < len && expr[i+1] >= '0' && expr[i+1] <= '9')) {
                var num = '';
                while (i < len && ((expr[i] >= '0' && expr[i] <= '9') || expr[i] === '.')) { num += expr[i]; i++; }
                tokens.push(num);
            }
            // Cell ref or function name
            else if ((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z')) {
                var word = '';
                while (i < len && ((expr[i] >= 'A' && expr[i] <= 'Z') || (expr[i] >= 'a' && expr[i] <= 'z') || (expr[i] >= '0' && expr[i] <= '9'))) {
                    word += expr[i]; i++;
                }
                // Check if it's a function call
                if (i < len && expr[i] === '(') {
                    var depth = 1; var inner = ''; i++; // skip (
                    while (i < len && depth > 0) {
                        if (expr[i] === '(') depth++;
                        else if (expr[i] === ')') { depth--; if (depth === 0) { i++; break; } }
                        inner += expr[i]; i++;
                    }
                    tokens.push(word.toUpperCase() + '(' + inner + ')');
                } else {
                    tokens.push(word);
                }
            }
            // Operators
            else if (ch === '+' || ch === '-' || ch === '*' || ch === '/') {
                tokens.push(ch); i++;
            }
            // Comparison operators
            else if (ch === '>' || ch === '<' || ch === '=') {
                var op = ch; i++;
                if (i < len && (expr[i] === '=' || (ch === '<' && expr[i] === '>'))) { op += expr[i]; i++; }
                tokens.push(op);
            }
            else if (ch === '(' || ch === ')') {
                tokens.push(ch); i++;
            }
            else {
                i++; // skip unknown
            }
        }
        return tokens;
    }

    // Evaluate a flat list of tokens with operator precedence
    function evalTokenList(tokens, cellData, visited) {
        // First resolve all non-operator tokens to values
        var values = [];
        var ops = [];
        for (var i = 0; i < tokens.length; i++) {
            var t = tokens[i];
            if (t === '+' || t === '-' || t === '*' || t === '/' ||
                t === '>' || t === '<' || t === '>=' || t === '<=' || t === '=' || t === '<>') {
                ops.push(t);
            } else {
                values.push(evalToken(t, cellData, visited));
            }
        }
        if (values.length === 0) return 0;
        if (ops.length === 0) return values[0];

        // Apply * and / first (left to right)
        var i = 0;
        while (i < ops.length) {
            if (ops[i] === '*' || ops[i] === '/') {
                var a = Number(values[i]), b = Number(values[i+1]);
                if (isNaN(a) || isNaN(b)) { values.splice(i, 2, '#VALUE!'); ops.splice(i, 1); }
                else {
                    var r = ops[i] === '*' ? a * b : (b === 0 ? '#DIV/0!' : a / b);
                    values.splice(i, 2, r); ops.splice(i, 1);
                }
            } else { i++; }
        }
        // Apply + and -
        i = 0;
        while (i < ops.length) {
            if (ops[i] === '+' || ops[i] === '-') {
                var a = Number(values[i]), b = Number(values[i+1]);
                if (isNaN(a) || isNaN(b)) { values.splice(i, 2, '#VALUE!'); ops.splice(i, 1); }
                else {
                    var r = ops[i] === '+' ? a + b : a - b;
                    values.splice(i, 2, r); ops.splice(i, 1);
                }
            } else { i++; }
        }
        // Apply comparisons
        i = 0;
        while (i < ops.length) {
            var a = values[i], b = values[i+1];
            var result;
            switch(ops[i]) {
                case '>':  result = a > b; break;
                case '<':  result = a < b; break;
                case '>=': result = a >= b; break;
                case '<=': result = a <= b; break;
                case '=':  result = a == b; break;
                case '<>': result = a != b; break;
                default: result = false;
            }
            values.splice(i, 2, result); ops.splice(i, 1);
        }
        return values[0];
    }

    // Main formula evaluator entry point
    function evaluateFormula(formula, cellData, visited) {
        if (!formula || formula.charAt(0) !== '=') return formula;
        try {
            var expr = formula.substring(1).trim();
            return evalExpression(expr, cellData, visited);
        } catch(e) {
            return '#ERROR!';
        }
    }

    // Compute display value for a cell (evaluates formula if present)
    function getDisplayValue(cellData, row, col) {
        var rd = cellData[row] || cellData[String(row)];
        if (!rd) return '';
        var cell = rd[col] || rd[String(col)];
        if (!cell) return '';
        var raw = cell.f !== undefined ? cell.f : (cell.v !== undefined ? String(cell.v) : '');
        if (typeof raw === 'string' && raw.charAt(0) === '=') {
            var result = evaluateFormula(raw, cellData, {});
            if (typeof result === 'number') {
                // Round display to avoid floating point noise
                return String(Math.round(result * 1e10) / 1e10);
            }
            return String(result);
        }
        return String(raw);
    }

    // Get the raw (editable) value — formula string or plain value
    function getEditValue(cellData, row, col) {
        var rd = cellData[row] || cellData[String(row)];
        if (!rd) return '';
        var cell = rd[col] || rd[String(col)];
        if (!cell) return '';
        if (cell.f !== undefined) return cell.f;
        return cell.v !== undefined ? String(cell.v) : '';
    }

    // ── Row/Column manipulation ──
    // Track last selected cell (buttons steal focus from inputs)
    var _lastActiveRow = -1;
    var _lastActiveCol = -1;

    function getActiveCell() {
        // First check if an input is currently focused
        var active = document.activeElement;
        if (active && active.tagName === 'INPUT' && active.hasAttribute('data-row')) {
            return { row: parseInt(active.getAttribute('data-row')), col: parseInt(active.getAttribute('data-col')) };
        }
        // Fall back to last tracked cell
        if (_lastActiveRow >= 0 && _lastActiveCol >= 0) {
            return { row: _lastActiveRow, col: _lastActiveCol };
        }
        return null;
    }

    function addRowAtEnd() {
        var sheet = window.spreadsheetData.sheets[_firstSheetId];
        var cur = sheet._gridRows || 10;
        if (cur >= MAX_ROWS) { showWarning('Maximum ' + MAX_ROWS + ' rows reached.'); return; }
        sheet._gridRows = cur + 1;
        reRender();
    }

    function insertRowAbove() {
        var cell = getActiveCell();
        if (!cell) { showWarning('Select a cell first to insert a row above it.'); return; }
        var sheet = window.spreadsheetData.sheets[_firstSheetId];
        var cur = sheet._gridRows || 10;
        if (cur >= MAX_ROWS) { showWarning('Maximum ' + MAX_ROWS + ' rows reached.'); return; }
        var cellData = sheet.cellData;
        // Shift all rows from bottom up to make room at cell.row
        for (var r = cur; r > cell.row; r--) {
            if (cellData[r - 1]) {
                cellData[r] = cellData[r - 1];
            } else {
                delete cellData[r];
            }
        }
        cellData[cell.row] = {}; // empty new row
        sheet._gridRows = cur + 1;
        _lastActiveRow = -1; _lastActiveCol = -1;
        reRender();
    }

    function deleteRow() {
        var cell = getActiveCell();
        if (!cell) { showWarning('Select a cell in the row you want to delete.'); return; }
        var sheet = window.spreadsheetData.sheets[_firstSheetId];
        var cur = sheet._gridRows || 10;
        if (cur <= 1) { showWarning('Cannot delete the last row.'); return; }
        var cellData = sheet.cellData;
        // Shift rows up from cell.row
        for (var r = cell.row; r < cur - 1; r++) {
            if (cellData[r + 1]) {
                cellData[r] = cellData[r + 1];
            } else {
                delete cellData[r];
            }
        }
        delete cellData[cur - 1];
        sheet._gridRows = cur - 1;
        _lastActiveRow = -1; _lastActiveCol = -1;
        reRender();
    }

    function addColAtEnd() {
        var sheet = window.spreadsheetData.sheets[_firstSheetId];
        var cur = sheet._gridCols || 5;
        if (cur >= MAX_COLS) { showWarning('Maximum ' + MAX_COLS + ' columns reached.'); return; }
        sheet._gridCols = cur + 1;
        reRender();
    }

    function insertColLeft() {
        var cell = getActiveCell();
        if (!cell) { showWarning('Select a cell first to insert a column to its left.'); return; }
        var sheet = window.spreadsheetData.sheets[_firstSheetId];
        var cur = sheet._gridCols || 5;
        if (cur >= MAX_COLS) { showWarning('Maximum ' + MAX_COLS + ' columns reached.'); return; }
        var cellData = sheet.cellData;
        var rowKeys = Object.keys(cellData);
        for (var ri = 0; ri < rowKeys.length; ri++) {
            var rd = cellData[rowKeys[ri]];
            if (!rd) continue;
            // Shift columns right from end
            for (var c = cur; c > cell.col; c--) {
                if (rd[c - 1]) {
                    rd[c] = rd[c - 1];
                } else {
                    delete rd[c];
                }
            }
            delete rd[cell.col]; // empty new column cell
        }
        sheet._gridCols = cur + 1;
        _lastActiveRow = -1; _lastActiveCol = -1;
        reRender();
    }

    function deleteCol() {
        var cell = getActiveCell();
        if (!cell) { showWarning('Select a cell in the column you want to delete.'); return; }
        var sheet = window.spreadsheetData.sheets[_firstSheetId];
        var cur = sheet._gridCols || 5;
        if (cur <= 1) { showWarning('Cannot delete the last column.'); return; }
        var cellData = sheet.cellData;
        var rowKeys = Object.keys(cellData);
        for (var ri = 0; ri < rowKeys.length; ri++) {
            var rd = cellData[rowKeys[ri]];
            if (!rd) continue;
            // Shift columns left from cell.col
            for (var c = cell.col; c < cur - 1; c++) {
                if (rd[c + 1]) {
                    rd[c] = rd[c + 1];
                } else {
                    delete rd[c];
                }
            }
            delete rd[cur - 1];
        }
        sheet._gridCols = cur - 1;
        _lastActiveRow = -1; _lastActiveCol = -1;
        reRender();
    }

    function reRender() {
        renderSpreadsheet(window.spreadsheetData);
        syncState();
    }

    // ── Rendering ──
    function renderSpreadsheet(data) {
        var editor = document.getElementById('spreadsheet-editor');
        if (!editor) return;

        if (!data.sheets) { showError('No sheets found. Keys: ' + Object.keys(data).join(', ')); return; }

        _firstSheetId = (data.sheetOrder && data.sheetOrder[0]) || Object.keys(data.sheets)[0];
        var sheet = data.sheets[_firstSheetId];
        if (!sheet || !sheet.cellData) { showError('No cellData in sheet: ' + _firstSheetId); return; }

        var cellData = sheet.cellData;
        var rowKeys = Object.keys(cellData).map(Number).sort(function(a,b){ return a-b; });

        var maxRow = rowKeys.length > 0 ? Math.max.apply(null, rowKeys) : 0;
        var maxCol = 0;
        rowKeys.forEach(function(r) {
            var rd = cellData[r] || cellData[String(r)];
            if (rd) {
                var cols = Object.keys(rd).map(Number);
                if (cols.length > 0) { var m = Math.max.apply(null, cols); if (m > maxCol) maxCol = m; }
            }
        });

        // Compute grid size from data. Use stored sheet dimensions if available,
        // otherwise default to at least 10 rows x 5 cols.
        var dataRows = Math.max(maxRow + 1, 10);
        var dataCols = Math.max(maxCol + 1, 5);
        // If sheet has explicit rowCount/columnCount from a previous add/delete, use those
        if (sheet._gridRows && sheet._gridRows > dataRows) dataRows = sheet._gridRows;
        if (sheet._gridCols && sheet._gridCols > dataCols) dataCols = sheet._gridCols;
        var displayRows = Math.min(dataRows, MAX_ROWS);
        var displayCols = Math.min(dataCols, MAX_COLS);

        // Store grid dimensions on sheet for persistence across re-renders
        if (!sheet._gridRows || sheet._gridRows < dataRows) sheet._gridRows = dataRows;
        if (!sheet._gridCols || sheet._gridCols < dataCols) sheet._gridCols = dataCols;
        var displayRows = Math.min(sheet._gridRows, MAX_ROWS);
        var displayCols = Math.min(sheet._gridCols, MAX_COLS);

        // Toolbar for row/column operations
        var toolbar = '<div id="grid-toolbar" style="display:flex;gap:6px;margin-bottom:8px;flex-wrap:wrap;">'
            + '<button type="button" id="btn-add-row" style="' + _btnStyle + '" title="Add row at bottom">+ Row</button>'
            + '<button type="button" id="btn-insert-row" style="' + _btnStyle + '" title="Insert row above selected cell">↑ Insert Row</button>'
            + '<button type="button" id="btn-del-row" style="' + _btnStyleDanger + '" title="Delete selected row">− Row</button>'
            + '<span style="border-left:1px solid #d0d7de;margin:0 4px;"></span>'
            + '<button type="button" id="btn-add-col" style="' + _btnStyle + '" title="Add column at right">+ Col</button>'
            + '<button type="button" id="btn-insert-col" style="' + _btnStyle + '" title="Insert column left of selected cell">← Insert Col</button>'
            + '<button type="button" id="btn-del-col" style="' + _btnStyleDanger + '" title="Delete selected column">− Col</button>'
            + '<span style="flex:1;"></span>'
            + '<span style="font-size:12px;color:#57606a;align-self:center;">' + displayRows + ' × ' + displayCols + '</span>'
            + '</div>';

        // Formula bar
        var formulaBar = '<div id="formula-bar" style="display:flex;align-items:center;gap:8px;margin-bottom:8px;padding:6px 8px;background:#f0f3f6;border-radius:4px;font-size:13px;">'
            + '<span id="cell-ref" style="font-weight:600;min-width:36px;color:#57606a;">A1</span>'
            + '<span style="color:#d0d7de;">|</span>'
            + '<span id="formula-display" style="font-family:monospace;color:#24292f;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">&nbsp;</span>'
            + '</div>';

        var html = '<table>';
        html += '<thead><tr><th></th>';
        for (var col = 0; col < displayCols; col++) {
            html += '<th>' + indexToCol(col) + '</th>';
        }
        html += '</tr></thead><tbody>';

        for (var row = 0; row < displayRows; row++) {
            html += '<tr><th>' + (row + 1) + '</th>';
            for (var c = 0; c < displayCols; c++) {
                var display = getDisplayValue(cellData, row, c);
                var edit = getEditValue(cellData, row, c);
                var isFormula = edit.charAt && edit.charAt(0) === '=';
                var tdClass = isFormula ? ' class="formula-cell"' : '';
                html += '<td' + tdClass + '><input type="text" name="cell-' + row + '-' + c + '"'
                    + ' value="' + escapeAttr(display) + '"'
                    + ' data-formula="' + escapeAttr(edit) + '"'
                    + ' data-row="' + row + '" data-col="' + c + '"'
                    + (isFormula ? ' title="' + escapeAttr(edit) + '"' : '')
                    + '></td>';
            }
            html += '</tr>';
        }
        html += '</tbody></table>';

        editor.innerHTML = toolbar + formulaBar + html;

        // Wire up toolbar buttons using mousedown to capture active cell before focus shifts
        document.getElementById('btn-add-row').addEventListener('mousedown', function(e) { e.preventDefault(); addRowAtEnd(); });
        document.getElementById('btn-insert-row').addEventListener('mousedown', function(e) { e.preventDefault(); insertRowAbove(); });
        document.getElementById('btn-del-row').addEventListener('mousedown', function(e) { e.preventDefault(); deleteRow(); });
        document.getElementById('btn-add-col').addEventListener('mousedown', function(e) { e.preventDefault(); addColAtEnd(); });
        document.getElementById('btn-insert-col').addEventListener('mousedown', function(e) { e.preventDefault(); insertColLeft(); });
        document.getElementById('btn-del-col').addEventListener('mousedown', function(e) { e.preventDefault(); deleteCol(); });

        // Wire up cell interactions
        var inputs = editor.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            // On focus: show formula in input, update formula bar, track active cell.
            // Default to "selected" mode (Excel: typing replaces, arrows navigate).
            inputs[i].addEventListener('focus', function() {
                var formula = this.getAttribute('data-formula') || '';
                if (formula && formula.charAt(0) === '=') {
                    this.value = formula;
                }
                var r = parseInt(this.getAttribute('data-row'));
                _lastActiveRow = r;
                var co = parseInt(this.getAttribute('data-col'));
                _lastActiveCol = co;
                var refEl = document.getElementById('cell-ref');
                var dispEl = document.getElementById('formula-display');
                if (refEl) refEl.textContent = indexToCol(co) + (r + 1);
                if (dispEl) dispEl.textContent = formula || this.value || '\u00a0';
                clearRefHighlight();
                _refLive = false;
                // Selected mode: highlight whole value so typing replaces.
                // Skip select() if focus came from a ref-pick navigation (handled there).
                if (!_editMode) {
                    var self = this;
                    // Defer to next tick so click positioning doesn't immediately deselect
                    setTimeout(function() {
                        if (document.activeElement === self) self.select();
                    }, 0);
                }
            });

            // Double-click enters edit mode (Excel: F2 / dbl-click puts cursor in cell)
            inputs[i].addEventListener('dblclick', function() {
                _editMode = true;
                var len = this.value.length;
                try { this.setSelectionRange(len, len); } catch(e) {}
            });

            // On blur: evaluate formula, show result
            inputs[i].addEventListener('blur', function() {
                var val = this.value;
                var r = parseInt(this.getAttribute('data-row'));
                var co = parseInt(this.getAttribute('data-col'));

                // Validate length
                if (val.length > MAX_CELL_LENGTH) {
                    val = val.substring(0, MAX_CELL_LENGTH);
                    showWarning('Cell content truncated to ' + MAX_CELL_LENGTH + ' characters.');
                }

                // Update in-memory data
                if (!window.spreadsheetData.sheets[_firstSheetId].cellData[r]) {
                    window.spreadsheetData.sheets[_firstSheetId].cellData[r] = {};
                }
                if (val === '') {
                    delete window.spreadsheetData.sheets[_firstSheetId].cellData[r][co];
                } else if (val.charAt(0) === '=') {
                    // Store as formula
                    window.spreadsheetData.sheets[_firstSheetId].cellData[r][co] = { f: val, v: '' };
                } else {
                    window.spreadsheetData.sheets[_firstSheetId].cellData[r][co] = { v: val };
                }

                // Re-evaluate all formula cells, then sync to hidden state input
                recalcAll();
                syncState();
            });

            // On change (fallback for form submission)
            inputs[i].addEventListener('change', function() {
                var val = this.value;
                var r = parseInt(this.getAttribute('data-row'));
                var co = parseInt(this.getAttribute('data-col'));
                if (!window.spreadsheetData.sheets[_firstSheetId].cellData[r]) {
                    window.spreadsheetData.sheets[_firstSheetId].cellData[r] = {};
                }
                if (val.charAt(0) === '=') {
                    window.spreadsheetData.sheets[_firstSheetId].cellData[r][co] = { f: val, v: '' };
                } else if (val !== '') {
                    window.spreadsheetData.sheets[_firstSheetId].cellData[r][co] = { v: val };
                }
                syncState();
            });
        }

        // Editor-level listeners (mouse + keyboard) attach once per editor element.
        // renderSpreadsheet() rebuilds the table on every reRender, so guard against
        // double-binding by stamping the attribute and reading it on subsequent calls.
        if (!editor.hasAttribute('data-listeners-installed')) {
            editor.setAttribute('data-listeners-installed', '1');
            editor.addEventListener('mousedown', onEditorMousedown);
            editor.addEventListener('keydown', onEditorKeydown);
        }
    }

    // ── Reference-target highlight ──
    function clearRefHighlight() {
        var prev = document.querySelectorAll('#spreadsheet-editor td.ref-target');
        for (var i = 0; i < prev.length; i++) prev[i].classList.remove('ref-target');
    }
    function setRefHighlight(row, col) {
        clearRefHighlight();
        var sel = '#spreadsheet-editor input[data-row="' + row + '"][data-col="' + col + '"]';
        var el = document.querySelector(sel);
        if (el && el.parentElement) el.parentElement.classList.add('ref-target');
    }

    // Insert (or replace if a live ref exists) a cell reference into the active input.
    // After this call, _refLive is true and _refRow/_refCol/_refStart/_refEnd describe
    // the live ref. Subsequent ref insertions REPLACE this slice.
    function setLiveRef(active, row, col) {
        if (row < 0 || col < 0) return;
        var ref = indexToCol(col) + (row + 1);
        var v = active.value;

        if (_refLive && _refStart >= 0 && _refEnd >= _refStart) {
            // Replace the previous live ref with the new ref text in the same slice.
            v = v.slice(0, _refStart) + ref + v.slice(_refEnd);
            _refEnd = _refStart + ref.length;
        } else {
            // Fresh insert at cursor.
            var s = typeof active.selectionStart === 'number' ? active.selectionStart : v.length;
            var e = typeof active.selectionEnd === 'number' ? active.selectionEnd : v.length;
            v = v.slice(0, s) + ref + v.slice(e);
            _refStart = s;
            _refEnd = s + ref.length;
        }

        active.value = v;
        active.setSelectionRange(_refEnd, _refEnd);
        _refLive = true;
        _refRow = row;
        _refCol = col;
        setRefHighlight(row, col);

        var dispEl = document.getElementById('formula-display');
        if (dispEl) dispEl.textContent = active.value;
    }

    // Commit the live ref so the next ref-pick acts as a fresh insert at the cursor.
    function commitLiveRef() {
        _refLive = false;
        _refStart = -1;
        _refEnd = -1;
        clearRefHighlight();
    }

    // ── Mouse handling: ref-pick during formula entry, otherwise enter selected mode ──
    function onEditorMousedown(e) {
        var active = document.activeElement;

        // Find the cell input under the click (if any)
        var target = e.target;
        while (target && target.tagName !== 'INPUT') target = target.parentElement;
        var clickedCell = (target && target.hasAttribute && target.hasAttribute('data-row')) ? target : null;

        // Case 1: clicking a different cell while entering a formula → pick a reference
        if (clickedCell && active && active !== clickedCell &&
            active.tagName === 'INPUT' && active.hasAttribute('data-row')) {
            var formula = active.value;
            if (formula && formula.charAt(0) === '=') {
                e.preventDefault();
                var tr = parseInt(clickedCell.getAttribute('data-row'));
                var tc = parseInt(clickedCell.getAttribute('data-col'));
                if (isNaN(tr) || isNaN(tc)) return;
                setLiveRef(active, tr, tc);
                return;
            }
        }

        // Case 2: clicking a cell normally → ensure we land in selected mode (not edit)
        // unless the user clicks the SAME cell that's already focused.
        if (clickedCell && clickedCell !== active) {
            _editMode = false;
        }
    }

    // ── Excel-style keyboard navigation ──
    function onEditorKeydown(e) {
        var active = document.activeElement;
        if (!active || active.tagName !== 'INPUT' || !active.hasAttribute('data-row')) return;

        var key = e.key;
        var shift = e.shiftKey;
        var r = parseInt(active.getAttribute('data-row'));
        var c = parseInt(active.getAttribute('data-col'));
        if (isNaN(r) || isNaN(c)) return;

        var inFormula = active.value.charAt(0) === '=';
        var isArrow = key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight';

        // Any non-arrow, non-modifier key in formula mode commits the live ref so the
        // next arrow/click inserts a fresh ref instead of replacing the last one.
        if (inFormula && _refLive && !isArrow && key !== 'Shift' && key !== 'Control' && key !== 'Alt' && key !== 'Meta') {
            commitLiveRef();
        }

        // F2: enter edit mode (cursor at end, no selection)
        if (key === 'F2') {
            e.preventDefault();
            _editMode = true;
            var len = active.value.length;
            try { active.setSelectionRange(len, len); } catch(err) {}
            return;
        }

        if (key === 'Enter' || key === 'Tab') {
            e.preventDefault();
            commitLiveRef();
            var sheet = window.spreadsheetData.sheets[_firstSheetId];
            var maxRow = (sheet._gridRows || 1) - 1;
            var maxCol = (sheet._gridCols || 1) - 1;
            var nr = r, nc = c;
            if (key === 'Enter') nr = shift ? Math.max(0, r - 1) : Math.min(maxRow, r + 1);
            else                 nc = shift ? Math.max(0, c - 1) : Math.min(maxCol, c + 1);
            if (nr !== r || nc !== c) {
                active.blur();
                focusCellAt(nr, nc, /* selectMode */ true);
            } else {
                // No room to move — commit current edit by re-selecting in place.
                active.blur();
                focusCellAt(r, c, /* selectMode */ true);
            }
            return;
        }

        if (key === 'Escape') {
            e.preventDefault();
            commitLiveRef();
            var cellData = window.spreadsheetData.sheets[_firstSheetId].cellData;
            var saved = getEditValue(cellData, r, c);
            active.value = saved;
            _editMode = false;
            active.blur();
            return;
        }

        if (key === 'Delete' || key === 'Backspace') {
            // In selected mode (whole-cell highlight), Delete clears the cell.
            // In edit mode, default behavior (delete one char) applies.
            if (!_editMode && active.selectionStart === 0 && active.selectionEnd === active.value.length) {
                e.preventDefault();
                active.value = '';
                // Don't enter edit mode — the cell stays in selected mode with empty value
                return;
            }
            return;
        }

        if (isArrow) {
            // Formula mode: arrows pick / move the ref target.
            if (inFormula && _editMode) {
                var dr = key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0;
                var dc = key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0;
                var baseR = _refLive ? _refRow : r;
                var baseC = _refLive ? _refCol : c;
                var tr = baseR + dr;
                var tc = baseC + dc;
                if (tr < 0 || tc < 0) return;
                e.preventDefault();
                setLiveRef(active, tr, tc);
                return;
            }

            // Edit mode (non-formula): only navigate at edges so user can position cursor.
            if (_editMode) {
                if (key === 'ArrowLeft' && active.selectionStart > 0) return;
                if (key === 'ArrowRight' && active.selectionEnd < active.value.length) return;
            }

            // Selected mode OR edit mode at edge: navigate cells, clamped to grid bounds.
            var sheet = window.spreadsheetData.sheets[_firstSheetId];
            var maxRow = (sheet._gridRows || 1) - 1;
            var maxCol = (sheet._gridCols || 1) - 1;
            var nr = r, nc = c;
            if (key === 'ArrowUp') nr = Math.max(0, r - 1);
            else if (key === 'ArrowDown') nr = Math.min(maxRow, r + 1);
            else if (key === 'ArrowLeft') nc = Math.max(0, c - 1);
            else if (key === 'ArrowRight') nc = Math.min(maxCol, c + 1);

            // Always swallow the arrow so it doesn't bubble to the dialog/page;
            // only blur+refocus when the target cell actually changes.
            e.preventDefault();
            if (nr !== r || nc !== c) {
                active.blur();
                focusCellAt(nr, nc, /* selectMode */ true);
            }
            return;
        }

        // Any printable single character in selected mode → enter edit mode replacing content.
        // (Excel: typing in a selected cell replaces its content.)
        if (!_editMode && key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            // Clear the value first; let the keystroke be appended naturally by the browser.
            active.value = '';
            _editMode = true;
            // Fall through — don't preventDefault, so the character is typed normally.
            return;
        }
    }

    // Focus a cell. selectMode = true ⇒ highlight whole content (Excel "selected" cell).
    // selectMode = false ⇒ cursor at end (Excel edit mode).
    function focusCellAt(row, col, selectMode) {
        if (row < 0 || col < 0) return;
        var sel = '#spreadsheet-editor input[data-row="' + row + '"][data-col="' + col + '"]';
        var el = document.querySelector(sel);
        if (!el) return;
        _editMode = !selectMode;
        el.focus();
        try {
            if (selectMode) {
                el.select();
            } else {
                var len = el.value.length;
                el.setSelectionRange(len, len);
            }
        } catch(e) { /* non-text input */ }
    }

    // Recalculate all cells and update display
    function recalcAll() {
        if (!_firstSheetId || !window.spreadsheetData) return;
        var cellData = window.spreadsheetData.sheets[_firstSheetId].cellData;
        var inputs = document.querySelectorAll('#spreadsheet-editor input');
        for (var i = 0; i < inputs.length; i++) {
            var inp = inputs[i];
            var r = parseInt(inp.getAttribute('data-row'));
            var c = parseInt(inp.getAttribute('data-col'));
            var edit = getEditValue(cellData, r, c);
            inp.setAttribute('data-formula', edit);
            // Only update display if not currently focused
            if (document.activeElement !== inp) {
                var display = getDisplayValue(cellData, r, c);
                inp.value = display;
                var td = inp.parentElement;
                if (td) {
                    if (edit.charAt && edit.charAt(0) === '=') {
                        td.className = 'formula-cell';
                        inp.title = edit;
                    } else {
                        td.className = '';
                        inp.title = '';
                    }
                }
            }
        }
    }

    function escapeAttr(str) {
        return str.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function escapeHtml(str) {
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }
})();
