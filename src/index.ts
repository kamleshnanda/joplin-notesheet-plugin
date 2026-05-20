import joplin from 'api';

const LOG = '[Notesheet]';

// Notesheet — a spreadsheet plugin for Joplin built on the Univer SDK.
// This is the M0 husk: the v0 popup model has been removed and the new
// inline architecture (Custom Editor + Univer engine) lands in M1.
joplin.plugins.register({
	onStart: async function () {
		console.info(LOG, 'plugin loaded (v1 scaffolding — no functionality yet)');
	},
});
