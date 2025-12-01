import globals from 'globals';
import tsParser from '@typescript-eslint/parser';
import tsPlugin from '@typescript-eslint/eslint-plugin';
import prettierPlugin from 'eslint-plugin-prettier';
import playwrightPlugin from 'eslint-plugin-playwright';
import sortKeysFixPlugin from 'eslint-plugin-sort-keys-fix';

export default [
  {
    // Use "ignores" to exclude non-TypeScript and generated files
    ignores: [
      'dist/**',
      'playwright-report/**',
      '.prettierrc.cjs',
      'eslint.config.mjs', // Exclude the config file itself
      'node_modules/**',
    ],
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
      parser: tsParser,
      parserOptions: {
        project: './tsconfig.json',
      },
      globals: {
        ...globals.browser,
        ...globals.node,
        ...globals.es2021,
      },
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
      prettier: prettierPlugin,
      playwright: playwrightPlugin,
      'sort-keys-fix': sortKeysFixPlugin,
    },
    rules: {
      'prettier/prettier': 'error',
      '@typescript-eslint/no-unused-vars': ['warn', { argsIgnorePattern: '^_' }],
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/explicit-module-boundary-types': 'off',
      'sort-keys': ['warn', 'asc', { caseSensitive: true, natural: false }],
      'playwright/no-wait-for-timeout': 'error',
    },
  },
  {
    files: ['*.spec.ts'],
    rules: {
      '@typescript-eslint/no-unused-vars': 'off',
      'playwright/no-wait-for-timeout': 'warn',
    },
  },
];
