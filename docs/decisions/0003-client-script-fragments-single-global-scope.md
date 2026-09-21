# 0003 — The client script is split into fragments that share one global scope

## Context

`WebEditor.js.html` had grown to about 2,300 lines mixing several concerns (image cache, list view, `*-table` widgets, edit view, export and move overlays). The client is plain classic `<script>` code: top-level `let`/`const`/`function` are shared globals, there is no bundler and no module system, and HtmlService includes files by pasting them into the page with `<?!= ... ?>` scriptlets. Small single-purpose files are cheaper for an agent (and a person) to navigate.

## Decision

Split it into fragments, one `<script>` block per file, by view/responsibility: `WebEditor.js.html` (core: bootstrap, data loading, shared state/helpers), `WebEditor.list.js.html`, `WebEditor.images.js.html`, `WebEditor.tables.js.html`, `WebEditor.edit.js.html`, `WebEditor.export.js.html`, `WebEditor.move.js.html`. They are included consecutively from `WebEditor.html`; that include order is the load order.

This is deliberately **not** a modularization. Separate classic scripts still share the global lexical scope, and the state is coupled (`schema` is read from every area), so the split changes navigation, not architecture. Rules that keep it safe:

- **Nothing runs at load time** except the host-sizing calls and the `DOMContentLoaded` listener (both in the core fragment). A top-level statement that reads another fragment's `const`/`let` would hit a `ReferenceError` if that fragment loads later. `tests/client-load.test.js` runs the fragments in include order in a `vm` with browser stubs to catch exactly that.
- **Each fragment owns its state** (`let`/`const` at its top); functions may freely call across fragments since they only run after load.
- **Tooling reads one list.** ESLint lints the fragments as one concatenated unit ([0002](0002-lint-client-fragments-with-inline-processor.md)); `tests/load.js` searches all fragments; `tests/html-contract.test.js` fails if a `WebEditor*.js.html` file isn't included.
- The split was made by a script that moved whole top-level statements (with their comments) unchanged, and verified afterwards: all 119 statements byte-identical and the multiset of non-blank lines equal to the original, apart from a two-line header comment per fragment.

## Consequences

- `git blame` for moved code starts at the split commit.
- The dialog itself still can't be exercised by any automated check here; after changing the fragment layout, smoke-test the editor by hand (open, filter, edit and save, export, move).
- If a fragment grows past a few hundred lines with more than one responsibility, split it the same way and add its include to `WebEditor.html`.
