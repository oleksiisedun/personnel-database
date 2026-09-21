# 0002 — Lint the client fragments as one unit through an inline ESLint processor

## Context

The web editor's client logic lives in `<script>` blocks inside `WebEditor*.js.html` files, which `tsc` cannot see. Everything there is a global, so a mistyped function name only fails at runtime, in the dialog, against real data. `eslint-plugin-html` is the usual answer but does not support ESLint 10. Since [0003](0003-client-script-fragments-single-global-scope.md) the script is split into several fragments that share one global scope, so linting each file on its own would flag every cross-fragment reference as undefined.

## Decision

`eslint.config.mjs` defines a small processor, applied to the entry file `src/WebEditor.js.html` with `no-undef` **on**. It reads the fragment list and order from the `<?!= ...createHtmlOutputFromFile('X.js') ?>` includes in `WebEditor.html` (via `tests/client-scripts.js`, the same module the tests use, so load order, lint order and test order can't drift), concatenates every fragment's `<script>` body, and lints the result as one file. `postprocess` maps each message back to its real fragment and line and tags it, e.g. `[WebEditor.edit.js.html] 'x' is not defined`. Concatenation also lets `no-redeclare` catch a function declared in two fragments.

`no-undef` stays **off** for the server `.js` files: Apps Script shares one global scope across files and `tsc`'s `checkJs` already reports genuinely undefined names there.

The processor assumes each fragment is exactly one `<script>` block and contains no scriptlets. Globals declared outside them (for example `INITIAL_MODE`, set by a scriptlet in `WebEditor.html`) must be listed in the config's `globals`.

## Consequences

- No dependency on an html plugin, and reported lines are the real per-file lines.
- Editors that lint a non-entry fragment in isolation get no result (no config matches it); the entry file is the one place lint runs.
- A fragment missing from the `WebEditor.html` includes would be neither loaded nor linted; `tests/html-contract.test.js` fails in that case.
