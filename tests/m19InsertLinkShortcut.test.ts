// M19 feature-2 (BACKLOG E4) — a dedicated shortcut opens Univer's insert-link
// panel while a Notesheet note is active.
//
// Why not Cmd/Ctrl+K: Joplin's default keymap binds it to `textLink`
// (laurent22/joplin @ v3.7.21 e41516e6, packages/lib/services/KeymapService.ts
// L39 darwin / L94 default), an Electron app-menu accelerator
// (packages/app-desktop/gui/MenuBar.tsx L732) enabled for plugin editors via
// `hasActivePluginEditor` (packages/app-desktop/gui/NoteEditor/
// editorCommandDeclarations.ts L24-28). Live (macOS) it only reaches Univer
// when the grid has focus; while editing a cell or with focus elsewhere it
// goes to Joplin's menu and nothing opens. So the plugin registers its own
// command + accelerator instead and must never touch `textLink`.
//
// `CmdOrCtrl+Alt+K` is unbound in both default keymap tables at v3.7.21
// (darwin `Option+Cmd+K`, default `Ctrl+Alt+K`) and in MenuBar.tsx's literal
// accelerators. The Univer side is pinned below against the installed sources.

import { existsSync, readFileSync } from 'fs';
import { dirname, join } from 'path';
import * as apiTypes from '../api/types';

// ── Fake `joplin` host API ────────────────────────────────────────────────
// tests/__mocks__/api.ts is a bare `{}`; this suite needs a recording fake.
// Both `api` and `api/types` map to that one mock file (jest.config.js), so a
// single factory supplies the default export and the real enums.

interface RegisteredCommand {
    name: string;
    label?: string;
    enabledCondition?: string;
    execute: (...args: unknown[]) => Promise<unknown>;
}

const fake = {
    commands: [] as RegisteredCommand[],
    menuItems: [] as { id: string; command: string; location: unknown; options?: unknown }[],
    posted: [] as { handle: string; message: unknown }[],
    active: new Map<string, boolean>(),
    visible: new Map<string, boolean>(),
    setups: [] as ((handle: string) => Promise<void>)[],
    onStart: null as null | (() => Promise<void>),
};

jest.mock('api', () => {
    const joplin = {
        plugins: {
            register: (p: { onStart: () => Promise<void> }) => {
                fake.onStart = p.onStart;
            },
        },
        contentScripts: { register: jest.fn(async () => undefined) },
        views: {
            editors: {
                register: jest.fn(
                    async (_id: string, cb: { onSetup: (h: string) => Promise<void> }) => {
                        fake.setups.push(cb.onSetup);
                    },
                ),
                setHtml: jest.fn(async () => ''),
                addScript: jest.fn(async () => undefined),
                onMessage: jest.fn(async () => undefined),
                onUpdate: jest.fn(async () => undefined),
                saveNote: jest.fn(async () => undefined),
                postMessage: jest.fn((handle: string, message: unknown) => {
                    fake.posted.push({ handle, message });
                }),
                isActive: jest.fn(async (handle: string) => {
                    if (!fake.active.has(handle)) throw new Error(`no view ${handle}`);
                    return fake.active.get(handle) as boolean;
                }),
                isVisible: jest.fn(async (handle: string) => {
                    if (!fake.visible.has(handle)) throw new Error(`no view ${handle}`);
                    return fake.visible.get(handle) as boolean;
                }),
            },
            toolbarButtons: { create: jest.fn(async () => undefined) },
            menuItems: {
                create: jest.fn(
                    async (id: string, command: string, location: unknown, options?: unknown) => {
                        fake.menuItems.push({ id, command, location, options });
                    },
                ),
            },
            dialogs: { showMessageBox: jest.fn(), showOpenDialog: jest.fn() },
        },
        commands: {
            register: jest.fn(async (c: RegisteredCommand) => {
                fake.commands.push(c);
            }),
            execute: jest.fn(async () => undefined),
        },
        workspace: { selectedNote: jest.fn(), selectedFolder: jest.fn() },
        data: { get: jest.fn(), post: jest.fn() },
    };

    const types = jest.requireActual('../api/types');
    return { __esModule: true, default: joplin, ...types };
});

async function startPlugin(): Promise<void> {
    jest.isolateModules(() => {
        require('../src/index');
    });
    if (!fake.onStart) throw new Error('plugin did not call joplin.plugins.register');
    await fake.onStart();
}

function insertLinkCommand(): RegisteredCommand {
    const c = fake.commands.find((x) => x.name === 'notesheetInsertLink');
    if (!c) throw new Error('notesheetInsertLink was not registered');
    return c;
}

/** Simulate Joplin creating one editor view per window (onSetup per handle). */
async function setupViews(...handles: string[]): Promise<void> {
    for (const h of handles) {
        for (const s of fake.setups) await s(h);
    }
}

describe('M19 E4 — notesheetInsertLink command registration', () => {
    beforeAll(async () => {
        await startPlugin();
    });

    beforeEach(() => {
        fake.posted.length = 0;
        fake.active.clear();
        fake.visible.clear();
    });

    test('onStart registers notesheetInsertLink gated on hasActivePluginEditor', () => {
        const c = insertLinkCommand();
        expect(c.label).toBe('Insert link in spreadsheet');
        expect(c.enabledCondition).toBe('hasActivePluginEditor');
    });

    test('a Tools menu item carries the verified CmdOrCtrl+Alt+K accelerator', () => {
        const items = fake.menuItems.filter((m) => m.command === 'notesheetInsertLink');
        expect(items).toHaveLength(1);
        expect(items[0].location).toBe(apiTypes.MenuItemLocation.Tools);
        expect(items[0].options).toEqual({ accelerator: 'CmdOrCtrl+Alt+K' });
    });

    test('no menu item reuses Cmd/Ctrl+K and nothing registers textLink', () => {
        const names = fake.commands.map((c) => c.name);
        expect(names).not.toContain('textLink');
        for (const m of fake.menuItems) {
            const acc = (m.options as { accelerator?: string } | undefined)?.accelerator;
            if (acc) expect(acc.replace(/\s/g, '')).not.toMatch(/^(CmdOrCtrl|Cmd|Ctrl)\+K$/i);
        }
    });

    test('running the command posts { type: "insertLink" } to the active, visible editor', async () => {
        await setupViews('view-1');
        fake.active.set('view-1', true);
        fake.visible.set('view-1', true);
        await insertLinkCommand().execute();
        expect(fake.posted).toEqual([{ handle: 'view-1', message: { type: 'insertLink' } }]);
    });

    test('posts nothing when isActive returns false', async () => {
        await setupViews('view-1');
        fake.active.set('view-1', false);
        fake.visible.set('view-1', false);
        await insertLinkCommand().execute();
        expect(fake.posted).toEqual([]);
    });

    test('posts nothing when the editor is active but hidden (user toggled to the markdown editor)', async () => {
        await setupViews('view-1');
        fake.active.set('view-1', true);
        fake.visible.set('view-1', false);
        await insertLinkCommand().execute();
        expect(fake.posted).toEqual([]);
    });

    test('with several views (second window) only the active one is targeted', async () => {
        await setupViews('view-1', 'view-2');
        fake.active.set('view-1', false);
        fake.visible.set('view-1', false);
        fake.active.set('view-2', true);
        fake.visible.set('view-2', true);
        await insertLinkCommand().execute();
        expect(fake.posted).toEqual([{ handle: 'view-2', message: { type: 'insertLink' } }]);
    });

    test('a view that no longer exists (isActive throws) is skipped, never posted to', async () => {
        await setupViews('gone', 'view-2');
        // 'gone' is absent from the maps, so the fake isActive throws for it.
        fake.active.set('view-2', true);
        fake.visible.set('view-2', true);
        await expect(insertLinkCommand().execute()).resolves.not.toThrow();
        expect(fake.posted).toEqual([{ handle: 'view-2', message: { type: 'insertLink' } }]);
    });

    test('fails closed: every known view is gone (isActive throws) → no post, no throw', async () => {
        const c = insertLinkCommand();
        // All handles from earlier tests are unknown to the fake now → isActive throws.
        await expect(c.execute()).resolves.not.toThrow();
        expect(fake.posted).toEqual([]);
    });
});

// ── Webview side: the message → Univer command bridge ─────────────────────
describe('M19 E4 — webview insert-link bridge', () => {
    const bridge = require('../src/editor/insertLink') as typeof import('../src/editor/insertLink');

    test('recognises only { type: "insertLink" }', () => {
        expect(bridge.isInsertLinkMessage({ type: 'insertLink' })).toBe(true);
        expect(bridge.isInsertLinkMessage({ type: 'load', snapshot: {} })).toBe(false);
        expect(bridge.isInsertLinkMessage(null)).toBe(false);
        expect(bridge.isInsertLinkMessage('insertLink')).toBe(false);
    });

    test('dispatches the injected Univer command id', () => {
        const calls: unknown[][] = [];
        const api = {
            executeCommand: (...args: unknown[]) => {
                calls.push(args);
                return Promise.resolve(true);
            },
        };
        expect(bridge.openUniverInsertLink(api, 'the.injected.id')).toBe(true);
        expect(calls).toEqual([['the.injected.id']]);
    });

    test('fails closed with no Univer instance, no executeCommand, or an empty id', () => {
        expect(bridge.openUniverInsertLink(null, 'x')).toBe(false);
        expect(bridge.openUniverInsertLink({}, 'x')).toBe(false);
        const api = { executeCommand: jest.fn(() => Promise.resolve(true)) };
        expect(bridge.openUniverInsertLink(api, '')).toBe(false);
        expect(api.executeCommand).not.toHaveBeenCalled();
    });

    test('a throwing or rejecting executeCommand never escapes', async () => {
        const throwing = {
            executeCommand: () => {
                throw new Error('boom');
            },
        };
        expect(bridge.openUniverInsertLink(throwing, 'x')).toBe(false);
        const rejecting = { executeCommand: () => Promise.reject(new Error('nope')) };
        expect(bridge.openUniverInsertLink(rejecting, 'x')).toBe(true);
        await new Promise((r) => setTimeout(r, 0)); // let the rejection be handled
    });
});

// ── Contract: pin the Univer constants editorView injects ─────────────────
// editorView passes `InsertLinkShortcut.id` (re-exported by the direct
// dependency @univerjs/preset-sheets-hyper-link) to the bridge. Read the
// installed sheets-hyper-link-ui build, resolved FROM that preset (not via npm
// hoisting), so an upgrade that renames the operation or rebinds the shortcut
// fails here instead of silently breaking the live shortcut.
function packageRoot(name: string, fromDir: string): string {
    let dir = dirname(require.resolve(name, { paths: [fromDir] }));
    for (;;) {
        const pj = join(dir, 'package.json');
        if (existsSync(pj) && JSON.parse(readFileSync(pj, 'utf8')).name === name) return dir;
        const up = dirname(dir);
        if (up === dir) throw new Error(`no package root for ${name}`);
        dir = up;
    }
}

const PRESET_ROOT = packageRoot('@univerjs/preset-sheets-hyper-link', __dirname);
const LINK_UI_ROOT = packageRoot('@univerjs/sheets-hyper-link-ui', PRESET_ROOT);
const LINK_UI_ES = readFileSync(join(LINK_UI_ROOT, 'lib', 'es', 'index.js'), 'utf8');

describe('M19 E4 — contract with installed @univerjs/sheets-hyper-link-ui', () => {
    test('the preset re-exports sheets-hyper-link-ui (so InsertLinkShortcut is importable)', () => {
        const presetEs = readFileSync(join(PRESET_ROOT, 'lib', 'es', 'index.js'), 'utf8');
        expect(presetEs).toMatch(/export \* from "@univerjs\/sheets-hyper-link-ui"/);
        expect(LINK_UI_ES).toMatch(/^export \{[^}]*\bInsertLinkShortcut\b[^}]*\};?$/m);
    });

    test("InsertLinkShortcut is Univer's Cmd/Ctrl+K and fires the insert-hyper-link-toolbar operation", () => {
        expect(LINK_UI_ES).toMatch(
            /const InsertLinkShortcut = \{\s*id: InsertHyperLinkToolbarOperation\.id,\s*binding: KeyCode\.K \| MetaKeys\.CTRL_COMMAND,/,
        );
        expect(LINK_UI_ES).toMatch(
            /const InsertHyperLinkToolbarOperation = \{\s*type: CommandType\.OPERATION,\s*id: "sheet\.operation\.insert-hyper-link-toolbar",/,
        );
    });

    test('no Univer shortcut in the installed tree binds K together with Alt', () => {
        // Walk @univerjs/* from the direct dependencies in package.json, each
        // resolved from its dependent, so the scan covers what webpack bundles.
        const pkg = JSON.parse(readFileSync(join(__dirname, '..', 'package.json'), 'utf8'));
        const seen = new Map<string, string>();
        const queue: [string, string][] = Object.keys(pkg.dependencies)
            .filter((n) => n.startsWith('@univerjs/'))
            .map((n) => [n, join(__dirname, '..')]);
        while (queue.length) {
            const [name, from] = queue.shift() as [string, string];
            if (seen.has(name)) continue;
            let root: string;
            try {
                root = packageRoot(name, from);
            } catch {
                continue; // types-only or optional package with no runtime entry
            }
            seen.set(name, root);
            const deps = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).dependencies;
            for (const d of Object.keys(deps ?? {})) {
                if (d.startsWith('@univerjs/')) queue.push([d, root]);
            }
        }
        expect(seen.has('@univerjs/sheets-hyper-link-ui')).toBe(true);
        const offenders: string[] = [];
        let kBindings = 0;
        for (const [name, root] of seen) {
            const es = join(root, 'lib', 'es', 'index.js');
            if (!existsSync(es)) continue;
            const src = readFileSync(es, 'utf8');
            for (const m of src.matchAll(/binding:\s*([^,}\n]+)/g)) {
                const expr = m[1];
                if (!/KeyCode\.K\b/.test(expr)) continue;
                kBindings++;
                if (/ALT/.test(expr)) offenders.push(`${name}: ${expr.trim()}`);
            }
        }
        expect(kBindings).toBeGreaterThan(0); // the scan really saw Univer's Cmd/Ctrl+K
        expect(offenders).toEqual([]);
    });
});
