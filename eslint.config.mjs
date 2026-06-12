// ESLint flat config (ESLint 9). Intentionally LENIENT in this first pass:
// the goal is to catch real correctness/safety issues and unused code, not
// to reformat 58 existing TS files or drown contributors in style nits.
// Prettier owns formatting (see .prettierrc.json). Tighten rules over time
// by promoting individual rules from 'warn' to 'error' as the codebase is
// cleaned, rather than enabling a strict preset all at once.
import js from '@eslint/js';
import tseslint from 'typescript-eslint';

export default tseslint.config(
    {
        // Don't lint build output, deps, or generated/vendored files.
        ignores: [
            'dist/**',
            'publish/**',
            'node_modules/**',
            'api/**', // Joplin-generated plugin API typings
            '**/*.js', // config + emitted JS; we lint TS/TSX sources
            '**/*.cjs',
            '**/*.mjs',
        ],
    },
    js.configs.recommended,
    ...tseslint.configs.recommended,
    {
        files: ['**/*.ts', '**/*.tsx'],
        rules: {
            // Real-bug guards stay as errors.
            'no-debugger': 'error',
            'no-cond-assign': 'error',
            'no-dupe-keys': 'error',
            'no-unreachable': 'error',
            // Pragmatic relaxations for an existing codebase using exceljs's
            // loosely-typed object model and Univer's `any`-heavy facade.
            // These start as warnings so they surface without blocking; the
            // CI lint step treats warnings as non-fatal until we ratchet up.
            '@typescript-eslint/no-explicit-any': 'off',
            '@typescript-eslint/no-unused-vars': [
                'warn',
                { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
            ],
            '@typescript-eslint/no-non-null-assertion': 'off',
            '@typescript-eslint/ban-ts-comment': 'warn',
            'no-empty': ['warn', { allowEmptyCatch: true }],
        },
    },
    {
        // Tests may use more permissive patterns (mocks, casts) and
        // legitimately embed irregular characters as fixture data in code
        // (e.g. MaliciousValues round-trip asserts BOM / bidi-override /
        // zero-width chars literally — editing those out would defeat the
        // test). Allow both here rather than scattering inline disables.
        files: ['tests/**/*.ts', 'tests/**/*.tsx'],
        rules: {
            '@typescript-eslint/no-require-imports': 'off',
            'no-irregular-whitespace': 'off',
            'no-control-regex': 'off',
        },
    },
    {
        // PGE harness + build scripts are Node CommonJS-style by nature;
        // require() is idiomatic there, not a mistake.
        files: ['scripts/**/*.ts'],
        rules: {
            '@typescript-eslint/no-require-imports': 'off',
        },
    },
);
