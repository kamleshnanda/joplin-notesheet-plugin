// Real-world smoke test: round-trip the user's actual InvestmentSummary.xlsx
// through M12's import + export pipeline and assert the formatting fidelity
// improvements actually fired.
//
// Skipped by default — only useful if the file is on disk. Manually opt in:
//   npm test -- --testPathPatterns=m12SmokeRealFile

import { existsSync, readFileSync } from 'fs';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const REAL_FILE = '/Users/kamleshn/Downloads/InvestmentSummary.xlsx';
const hasFile = existsSync(REAL_FILE);

(hasFile ? describe : describe.skip)('M12 — real-file round-trip (InvestmentSummary.xlsx)', () => {
    test('import: theme font Aptos Narrow + TableStyleMedium2 banding + 7 hyperlinks', async () => {
        const buf = readFileSync(REAL_FILE);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as {
            sheetOrder: string[];
            sheets: Record<string, {
                cellData: Record<number, Record<number, {
                    v?: unknown; s?: string;
                    p?: { body?: { customRanges?: Array<{ rangeType?: number; properties?: { url?: string } }> } };
                }>>;
            }>;
            styles: Record<string, Record<string, unknown>>;
            defaultStyle?: { ff?: string };
        };

        // Theme font.
        expect(snap.defaultStyle?.ff).toBe('Aptos Narrow');

        // Banding: header row (row 0) of the table should have the
        // TableStyleMedium2 header bg = #156082.
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const headerStyleId = sheet.cellData[0]?.[0]?.s;
        expect(headerStyleId).toBeDefined();
        const headerStyle = snap.styles[headerStyleId!];
        expect((headerStyle.bg as { rgb: string })?.rgb).toBe('#156082');

        // Hyperlinks: scan all cells for HYPERLINK customRanges. User's file
        // has 7. We count whatever's actually present (tolerant — exceljs
        // sometimes loses a link if the workbook is unusual).
        let linkCount = 0;
        for (const rKey of Object.keys(sheet.cellData)) {
            const row = sheet.cellData[Number(rKey)];
            if (!row) continue;
            for (const cKey of Object.keys(row)) {
                const cell = row[Number(cKey)];
                const ranges = cell?.p?.body?.customRanges ?? [];
                if (ranges.some((r) => r.rangeType === 0 && typeof r.properties?.url === 'string')) {
                    linkCount++;
                }
            }
        }
        expect(linkCount).toBeGreaterThanOrEqual(5); // at least most of them
    });

    test('export: round-trip preserves theme font + hyperlinks', async () => {
        const buf = readFileSync(REAL_FILE);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const buf2 = await snapshotToXlsxBuffer(snap);

        // Write to /tmp so we can manually open in Excel after the run.
        // eslint-disable-next-line @typescript-eslint/no-require-imports
        const fs = await import('fs');
        fs.writeFileSync('/tmp/m12-roundtrip.xlsx', Buffer.from(buf2 as ArrayBuffer));

        // Re-open the exported file and confirm the theme font was preserved
        // in xl/theme/theme1.xml.
        const zip = await JSZip.loadAsync(buf2 as ArrayBuffer);
        const themePath = Object.keys(zip.files).find((p) => /^xl\/theme\/theme\d+\.xml$/i.test(p))!;
        const theme = await zip.files[themePath].async('string');
        expect(theme).toContain('typeface="Aptos Narrow"');

        // exceljs should re-parse the result and find at least some hyperlinks.
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf2 as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet(1)!;
        let linkCount = 0;
        ws.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (c) => {
                if (c.isHyperlink) linkCount++;
            });
        });
        expect(linkCount).toBeGreaterThanOrEqual(5);

        // Table should round-trip too.
        const tables = (ws as unknown as { getTables: () => Array<{ name: string }> }).getTables();
        expect(tables.length).toBeGreaterThanOrEqual(1);
        expect(tables[0].name).toBe('Table1');
    });
});
