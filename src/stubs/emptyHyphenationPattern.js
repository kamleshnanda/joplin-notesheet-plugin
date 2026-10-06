// Stand-in for a Univer hyphenation pattern module that the build strips
// (see webpack.config.js → hyphenationPatternStub). Exports nothing, so
// engine-render's Hyphen.loadPattern() finds no pattern and returns early;
// every caller checks hasPattern() first, so that language just doesn't
// hyphenate — it never throws.
export {};
