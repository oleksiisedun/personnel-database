import js from '@eslint/js';
import globals from 'globals';
import path from 'node:path';
import { createRequire } from 'node:module';

const { clientScripts } = createRequire(import.meta.url)('./tests/client-scripts.js');

// The client script is split across several WebEditor.*.js.html fragments that
// share one global scope in the browser (see WebEditor.html). To keep `no-undef`
// meaningful they are linted as ONE unit: this processor runs on the entry file
// (src/WebEditor.js.html), concatenates every fragment's <script> body in include
// order, and maps each message back to its real file and line. Bodies are taken
// from the character after `<script>`, so body line N is file line N.
// (eslint-plugin-html does not support ESLint 10, and the scripts contain no scriptlets.)
const srcDir = path.join(import.meta.dirname, 'src');
const layouts = new Map(); // entry file name -> [{ file, startLine }] for postprocess()

const scriptBody = (html) => html.slice(html.indexOf('<script>') + '<script>'.length, html.lastIndexOf('</script>'));

const htmlScript = {
  processors: {
    script: {
      preprocess: (text, filename) => {
        const entry = path.basename(filename);
        const fragments = clientScripts(srcDir).map((f) => (f.file === entry ? { ...f, text } : f));
        let line = 0;
        const layout = [];
        const code = fragments.map(({ file, text: html }) => {
          const body = scriptBody(html);
          layout.push({ file, startLine: line });
          line += body.split('\n').length;
          return body;
        });
        layouts.set(filename, layout);
        return [code.join('\n')];
      },
      postprocess: (messages, filename) => {
        const layout = layouts.get(filename);
        return messages.flat().map((msg) => {
          const at = [...layout].reverse().find((l) => l.startLine < msg.line);
          return { ...msg, line: msg.line - at.startLine, endLine: msg.endLine && msg.endLine - at.startLine, message: `[${at.file}] ${msg.message}` };
        });
      },
      supportsAutofix: false,
    },
  },
};

// Correctness-only rules. Types are already covered by `npm run typecheck`
// (src/*.js files); the main gap this fills is the JS embedded in
// the WebEditor.*.js.html fragments, which tsc cannot see.
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
    rules: {
      ...rules,
      'no-undef': 'off',
      'no-unused-vars': 'off',
      // Opening a spreadsheet the user can't access fails the whole execution even
      // when caught; openSpreadsheetSafely() (Code.js) is the only allowed caller.
      'no-restricted-properties': ['error',
        { object: 'SpreadsheetApp', property: 'openById', message: 'Use openSpreadsheetSafely(id).' },
        { object: 'SpreadsheetApp', property: 'openByUrl', message: 'Use openSpreadsheetSafely() with parseDriveId(url).' },
      ],
    },
  },

  // Client-side script inside the HtmlService dialog (all WebEditor.*.js.html
  // fragments, linted as one unit via the entry file). Not a module: every
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
