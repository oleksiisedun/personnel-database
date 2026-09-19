import js from '@eslint/js';
import globals from 'globals';

// Extracts the <script> body from an .html file, padded with blank lines so
// reported line numbers match the real file. (eslint-plugin-html does not
// support ESLint 10, and this file has no scriptlets inside its <script>.)
const htmlScript = {
  processors: {
    script: {
      preprocess: (text) => {
        const start = text.indexOf('<script>');
        const end = text.lastIndexOf('</script>');
        if (start === -1 || end === -1) return [];
        const before = text.slice(0, start + '<script>'.length);
        const padding = '\n'.repeat(before.split('\n').length - 1);
        return [padding + text.slice(start + '<script>'.length, end)];
      },
      postprocess: (messages) => messages.flat(),
      supportsAutofix: false,
    },
  },
};

// Correctness-only rules. Types are already covered by `npm run typecheck`
// (src/*.js files); the main gap this fills is the JS embedded in
// WebEditor.js.html, which tsc cannot see.
const rules = {
  ...js.configs.recommended.rules,
  'no-var': 'error',
  'prefer-const': 'error',
  eqeqeq: ['error', 'always', { null: 'ignore' }],
  'no-shadow': 'error',
  'no-loop-func': 'error',
  // Apps Script's runtime/typings target ES2020; Error `cause` is ES2022.
  'preserve-caught-error': 'off',
  'no-unused-vars': ['error', { args: 'none', caughtErrors: 'none' }],
};

export default [
  { ignores: ['node_modules/**', 'samples/**'] },

  // Server-side Apps Script files share one global scope across files, so
  // cross-file references are legitimate; tsc's checkJs already flags
  // genuinely undefined names here.
  {
    files: ['src/*.js'],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: 'script',
      globals: { ...globals.es2021 },
    },
    rules: { ...rules, 'no-undef': 'off', 'no-unused-vars': 'off' },
  },

  // Client-side script inside the HtmlService dialog. Not a module: every
  // top-level function is a global, so no-undef is the check tsc can't do.
  {
    files: ['src/WebEditor.js.html'],
    plugins: { html: htmlScript },
    processor: 'html/script',
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: 'script',
      globals: {
        ...globals.browser,
        google: 'readonly',
        // Declared by a scriptlet in WebEditor.html, outside this <script>.
        INITIAL_MODE: 'readonly',
      },
    },
    rules: { ...rules, 'no-unused-vars': 'off' },
  },
];
