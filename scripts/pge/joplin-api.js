// Tiny CLI wrapper over Joplin's Web Clipper Data API.
//
// Usage:
//   node joplin-api.js create-note --title "Foo" --body-file body.md
//   node joplin-api.js list-notes [--folder-id <id>]
//   node joplin-api.js create-folder --title "PGE Smoke"
//   node joplin-api.js delete-note <id>
//
// The Data API is unauthenticated on localhost when Web Clipper is
// enabled with no token. If JOPLIN_TOKEN env var is set, it's
// appended as ?token=... on every request.

const fs = require('fs');

const API = process.env.JOPLIN_API || 'http://localhost:41184';
const TOKEN = process.env.JOPLIN_TOKEN || '';

function url(path, query = {}) {
    const q = new URLSearchParams(query);
    if (TOKEN) q.set('token', TOKEN);
    const qs = q.toString();
    return API + path + (qs ? '?' + qs : '');
}

async function request(method, path, body = null, query = {}) {
    const opts = {
        method,
        headers: { 'Content-Type': 'application/json' },
    };
    if (body !== null) opts.body = JSON.stringify(body);
    const res = await fetch(url(path, query), opts);
    if (!res.ok) {
        const text = await res.text();
        throw new Error(`${method} ${path} → ${res.status}: ${text}`);
    }
    return await res.json();
}

function getFlag(args, name) {
    const i = args.indexOf('--' + name);
    if (i === -1) return null;
    return args[i + 1];
}

async function main() {
    const [, , cmd, ...rest] = process.argv;
    if (!cmd) {
        console.error('usage: joplin-api.js <create-note|list-notes|create-folder|delete-note> [options]');
        process.exit(2);
    }

    if (cmd === 'create-note') {
        const title = getFlag(rest, 'title') || 'PGE evaluator note';
        const bodyFile = getFlag(rest, 'body-file');
        const folderId = getFlag(rest, 'folder-id');
        if (!bodyFile) throw new Error('--body-file is required');
        const body = fs.readFileSync(bodyFile, 'utf8');
        const payload = { title, body };
        if (folderId) payload.parent_id = folderId;
        const out = await request('POST', '/notes', payload);
        process.stdout.write(out.id + '\n');
        return;
    }

    if (cmd === 'list-notes') {
        const folderId = getFlag(rest, 'folder-id');
        const path = folderId ? `/folders/${folderId}/notes` : '/notes';
        const out = await request('GET', path, null, { fields: 'id,title' });
        for (const n of out.items ?? []) console.log(n.id, n.title);
        return;
    }

    if (cmd === 'create-folder') {
        const title = getFlag(rest, 'title') || 'PGE Smoke';
        const out = await request('POST', '/folders', { title });
        process.stdout.write(out.id + '\n');
        return;
    }

    if (cmd === 'delete-note') {
        const id = rest[0];
        if (!id) throw new Error('delete-note requires <id>');
        await request('DELETE', `/notes/${id}`);
        return;
    }

    console.error('unknown command: ' + cmd);
    process.exit(2);
}

main().catch((e) => { console.error(e.message); process.exit(1); });
