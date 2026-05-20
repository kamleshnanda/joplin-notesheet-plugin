/** @type {import('jest').Config} */
module.exports = {
    preset: 'ts-jest',
    testEnvironment: 'node',
    roots: ['<rootDir>/tests'],
    testMatch: ['**/*.test.ts'],
    moduleNameMapper: {
        // The Joplin plugin runtime injects an `api` global; for tests we
        // mock or ignore it. Tests should import only pure modules.
        '^api$': '<rootDir>/tests/__mocks__/api.ts',
        '^api/(.*)$': '<rootDir>/tests/__mocks__/api.ts',
    },
};
