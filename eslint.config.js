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
      '@rushstack/no-new-null': 'off'
    }
  }
];
