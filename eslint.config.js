const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    // The graph engine is the v1.x single-file application, extracted verbatim.
    // It is untyped by design (see src/engine/README.md) and carries @ts-nocheck;
    // linting it would report thousands of style findings against code whose
    // value is that it is UNCHANGED from the shipped, tested original.
    ignores: ['src/engine/engine.ts', 'src/engine/engineAssets.ts']
  },
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    },
    rules: {
      '@typescript-eslint/no-floating-promises': 'off',
      // SharePoint REST returns null (not undefined) for empty columns; the
      // DTO layer models that honestly.
      '@rushstack/no-new-null': 'off',
      // `void somePromise()` is how this codebase marks a deliberately un-awaited
      // async call (autosave flushes, poll ticks). Silence here would be worse than
      // the operator: the alternative is a bare call that reads like an oversight.
      'no-void': 'off',
      // A class referenced inside a method body of a class declared above it is safe:
      // the module is fully evaluated before any method runs. Functions and variables
      // stay checked, where the rule catches real temporal-dead-zone bugs.
      '@typescript-eslint/no-use-before-define': ['error', { classes: false }]
    }
  }
];
