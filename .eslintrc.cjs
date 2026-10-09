module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: ['eslint:recommended', 'plugin:@typescript-eslint/recommended'],
  env: {
    node: true,
    es2022: true,
  },
  ignorePatterns: ['dist/', 'coverage/', 'node_modules/'],
  parserOptions: {
    ecmaVersion: 2022,
    sourceType: 'module',
    warnOnUnsupportedTypeScriptVersion: false,
  },
  rules: {
    // The base rule cannot handle TypeScript types; delegate to the plugin rule below
    'no-unused-vars': 'off',
    // Existing financial handlers still use `any` (audit Finding 2.2, scheduled for Phase 3 Zod migration)
    '@typescript-eslint/no-explicit-any': 'off',
    // Non-null assertions are used pervasively for regex .match() results
    '@typescript-eslint/no-non-null-assertion': 'off',
    '@typescript-eslint/no-unused-vars': [
      'warn',
      { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
    ],
  },
};