## Deployment

```bash
# Push to the primary bound script project
clasp push

# Push to all target spreadsheets listed in clasp-targets.json
./clasp-push.sh   # or: npm run clasp-push
```

All deployable code (`*.js`, `WebEditor.*.html`, `appsscript.json`) lives in `src/`; `.clasp.json` (git-ignored) sets `"rootDir": "src"` so clasp pushes only that directory, and tooling/docs stay at the repo root. Paths below are relative to `src/` unless they name a root file. If `.clasp.json` is ever recreated (`clasp clone`/`create`), re-add `rootDir`.

There is no build step. Machine-checkable guardrails — run `npm run check` (all four, in sequence) after editing code. CI (`.github/workflows/check.yml`) runs the same command on every push to `main` and every pull request:

- `npm run typecheck` (`tsc -p jsconfig.json`, `checkJs` over `src/*.js` against `@types/google-apps-script`, non-strict). It can't see the JS embedded in the `WebEditor.*.js.html` client fragments. Since JSDoc is the only source of type info, a failure often means a stale `@param`/`@returns`.
- `npm run lint` (ESLint, `eslint.config.mjs`) — correctness-only rules over `src/*.js` and the `<script>` bodies of the `WebEditor.*.js.html` client fragments. A small inline processor in the config (`eslint-plugin-html` doesn't support ESLint 10) runs on the entry file `WebEditor.js.html` and concatenates every fragment in `WebEditor.html` include order into **one** lint unit, tagging each message with the real fragment file and line. Its main value is `no-undef` (plus `no-redeclare` for a function declared in two fragments) on the client code, the one place `tsc` can't reach: everything there is a global, so a typo'd function name otherwise only fails at runtime. `no-undef`/`no-unused-vars` are off for the server `.js` files (Apps Script shares one global scope across files; `tsc` covers undefined names there). Globals declared outside that `<script>` (e.g. `INITIAL_MODE`, set by a scriptlet in `WebEditor.html`) must be added to the config's `globals`. `preserve-caught-error` is off because `Error` `cause` needs ES2022 typings. `no-restricted-properties` bans `SpreadsheetApp.openById`/`openByUrl` in `src/*.js`; `openSpreadsheetSafely()` is the only exception (see below).
- `npm run lint:sh` (`shellcheck` + `shfmt -d` on `clasp-push.sh`; `.editorconfig` supplies shfmt's 2-space/`space_redirects` style). Needs `shellcheck` and `shfmt` installed locally; run `shfmt -w clasp-push.sh` to fix formatting.
- `npm test` (`node --test 'tests/*.test.js'`, no dependencies) — unit tests for the pure helpers in `Utils.js`/`Config.js`/`ExportValues.js`, a parity test keeping server `parseDriveId()` and client `extractDriveId()` aligned, and `tests/html-contract.test.js`, static checks for what only fails at runtime in the dialog: every literal `getElementById()` id exists in `WebEditor.html` (or is assigned in script), every `#toolbar` button has `btn-toolbar`, and every server function serving `WebEditor` sets `template.mode`, and every `WebEditor*.js.html` file is included by `WebEditor.html`. `tests/load.js` runs the Apps Script sources in a Node `vm` context (unmodified, no exports needed) and can lift one self-contained function out of any client fragment by its 2-space indent. `tests/client-load.test.js` runs the client fragments in include order in a `vm` with browser stubs, catching a parse error or a load-time reference to a later fragment (a `ReferenceError`). **Keep new logic in pure functions that take plain values** (as `normalizePhoneNumber()` does) so it stays testable without mocking `SpreadsheetApp`/`DriveApp`; add a test when you change one.

**Never run `clasp push`/`clasp-push.sh` or otherwise deploy/test changes yourself.** These scripts push live to real bound spreadsheets (including production personnel data across all targets in `clasp-targets.json`). Leave deployment and live testing to the user.

## File map

Google Apps Script project (V8 runtime) bound to a Google Spreadsheet; the web editor runs in an `HtmlService` modal opened from the Sheets menu.

- `Code.js` — menu (`onOpen`), modal openers (`openWebEditor`, `openPhotoExport`), `getSchemaAndData()`, row CRUD, image proxy, `fixPhoneNumbers`/`fixFullNames`
- `Utils.js` — shared helpers: spreadsheet/Handbook resolution, schema comparison, column lookups, Drive-ID parsing (`parseDriveId`), `*-table` cell codec (`_parseSubTable`/`_encodeSubTable`)
- `Export.js` — F-1/Wanted Card Docs export (`_exportDoc`), shared sheet-data loader, template placeholder/image helpers
- `ExportValues.js` — pure computed-value functions for the correspondence table (`_computeValue` and friends, `_calendarDuration`); no Apps Script services, unit-tested
- `ExportXlsx.js` — XLSX export (`exportXLSX`)
- `ExportPhotos.js` — S-КАДР photo export (`startPhotoExport`, `copyPhotosBatch`)
- `Import.js` — award import from S-КАДР
- `Config.js` — all constants
- `WebEditor.html` / `WebEditor.css.html` — client shell (its `<?!= ... ?>` include order is the client load order) and styles
- Client logic, split by view/responsibility into fragments that share one global scope (no modules; [ADR 0003](docs/decisions/0003-client-script-fragments-single-global-scope.md)): `WebEditor.js.html` (core: bootstrap `init()`, data loading/refresh, shared state and helpers), `WebEditor.list.js.html` (list view, filters, column panel), `WebEditor.images.js.html` (image cache, badges, lightbox, `extractDriveId`), `WebEditor.tables.js.html` (`*-table` mini table, table editor, codec), `WebEditor.edit.js.html` (edit view, validation, save/delete), `WebEditor.export.js.html` (export overlay flows), `WebEditor.move.js.html` (move overlay). Put new client code in the fragment for its view; a new fragment needs its include added to `WebEditor.html` (a test fails otherwise). Each fragment keeps its own state `let`/`const` at the top; nothing may run at load time except the host-sizing calls and the `DOMContentLoaded` listener in the core
- `tests/` (repo root) — `node --test` unit/contract tests and the `load.js` vm loader; not deployed
- `docs/decisions/` (repo root) — short ADRs for deliberate-but-surprising choices (see the end of the Architecture section)
- `samples/` (repo root) — example Database layout and F-1/Wanted Card templates for setting up a new deployment; not deployed

## Code conventions

JS style and JSDoc rules are in the global `~/.claude/CLAUDE.md`. Project-specific rule: all tuneable constants belong in `Config.js` (see the end of this file).

### CSS design system (`WebEditor.css.html`)

Concrete tokens and shared classes already defined here (see global CSS design-system conventions for the reuse-first principle):

- **Variables** (all colors are tokens — never write a raw hex outside `:root`; add a `--color-*` token instead): accent `--color-primary`/`-hover`/`-tint`, `--color-on-accent`, `--color-focus`/`-focus-shadow`; status `--color-danger`/`-hover`/`-tint`/`-shadow`, `--color-success`/`-tint`, `--color-warning`/`-hover`/`-tint`; text `--color-text`/`-muted`/`-subtle`/`-faint`; surfaces `--color-surface`/`-alt`/`-subtle`; borders `--color-border` (inputs), `--color-border-light` (table/panel lines), `--color-border-faint` (row separators); `--select-arrow` (dropdown chevron), `--radius`.
- **Shared classes**: `.btn-primary` (filled accent — Save/Move/Close), `.btn-secondary` (outline — Back/Cancel/toolbar), `.btn-danger` (outline, danger color — Delete, Discard Changes), `.btn-success` (outline, success color — Refresh), `.btn-warning` (outline, warning color — schema-mismatch ⚠ button, Reset), `.btn-dialog-action` (sizing only — `padding`/`font-size` for overlay action buttons: export Cancel/Proceed/Close, move Cancel/Confirm, schema-warning Close, unsaved-changes Keep Editing/Discard Changes — combined with a color class), `.btn-toolbar` (sizing only — `padding: 5px 12px` for every button inside `#toolbar` — combined with a color class the same way), `.overlay`/`.overlay-dialog` (modal scaffolding — export progress, move dialog, unsaved-changes dialog), `.dialog-actions` (`display:flex; gap:10px; justify-content:flex-end` action row; add `.dialog-actions--center` to center it — new overlay action rows should use this instead of a one-off `#foo-actions` ID rule), `.overlay-title`/`.overlay-text` (dialog heading and body copy — don't restyle titles/messages per ID or inline), `.log-list` (monospace scrollable per-item log box — used by `#move-log` and `#export-skip-log`; new per-item log/result lists should reuse this instead of a one-off `#foo-log` ID rule).
- **Global states**: `button:disabled` (`opacity: 0.5; cursor: default;`), text input/select focus (`border-color: var(--color-focus); box-shadow: 0 0 0 2px var(--color-focus-shadow);`).
- **Every new toolbar button must add `btn-toolbar` to its `class`**, alongside a color class (`class="btn-secondary btn-toolbar"`). Do not reintroduce a per-ID padding list — it's a duplication trap: a button (`#btn-export-xlsx`) once shipped without it and rendered with the browser's default padding.

## Architecture

### Sheet layout

The `Database` sheet has a fixed structure:
- Row 1 — column names (header)
- Row 2 — column types (`text`, `image`, `*-table`, `unit`, `origin`, `marital-status`, `sex`, `tin`, `number`, …)
- Row 3+ — one person record per row

The `Handbook` sheet holds schema metadata, dropdown/table-column definitions, the Master Mode toggle, the data folder cell, and export correspondence tables. `Config.js` defines each range address.

Every Handbook range read inside `getSchemaAndData()` (`getMasterMode()`, `getMasterSources()`, `getDataTypeOptionsMap()`, `getTableColumnsMap()`) is wrapped in try/catch with a safe fallback (`false`, `[]`, or `{}`), keeping the editor usable if a range is unreadable for some user. **New Handbook reads added to `getSchemaAndData()` must follow the same pattern.**

Handbook layout: single-value config lives in column A as a vertical list (a documentation label, then its value on the row(s) below), Master Mode sources in `B2:B`, and three open-ended tables from row 1 onward — export correspondence (`D2:F`), `*-table` sub-columns (`H2:J`), data types (`L2:AL`). Every address is a constant in `Config.js`; the cell-by-cell reference is [docs/spreadsheet-setup.md](docs/spreadsheet-setup.md) (don't restate it here). `A4` (`DATA_FOLDER`) is the shared "UNITS" folder, and each unit's own subfolder is resolved by `getUnitDataFolder()` (see [docs/architecture-master-mode.md](docs/architecture-master-mode.md)); the correspondence table is covered in [docs/architecture-export-import.md](docs/architecture-export-import.md) and data types in [docs/architecture-web-editor.md](docs/architecture-web-editor.md).

All Drive-ID/spreadsheet-link Handbook cells tolerate either a bare ID or a full sharing URL — every read site pipes the value through `parseDriveId()` (directly or via `getDriveIdFromHandbook()`).

Every target spreadsheet has its own local `Handbook` sheet. Ones that share config with a central source do so via `IMPORTRANGE` formulas in their own cells, not any code-level fallback — to the script every `Handbook` is self-contained.

### Startup healthcheck (`handbookCheck`)

The per-range try/catch fallbacks make a missing `Handbook` sheet fail silently, so `openWebEditor()` and `openPhotoExport()` (`Code.js`) call `assertHandbookHealthy(actionName)` before creating the modal. It runs `handbookCheck()` (`Utils.js`), which checks **only** that the `Database` and `Handbook` sheets exist by name — not cell contents or range structure, since most Handbook config is read defensively or is optional per-feature (a blank `A6`/`A7` disables the actual-personnel filter; a blank `EXPORT_FOLDER_CELL` throws its own clear error at export time). If either sheet is missing, it shows a `ui.alert()` and the modal is never opened.

### How the HTML template works

`WebEditor.html` is served via `HtmlService.createTemplateFromFile()` and includes CSS and JS with scriptlets:

```html
<?!= HtmlService.createHtmlOutputFromFile('WebEditor.css').getContent(); ?>
<?!= HtmlService.createHtmlOutputFromFile('WebEditor.js').getContent(); ?>
```

`<?!=` (with `!`) injects the content **without** sanitization, so the raw `<style>`/`<script>` tags are kept. Column width CSS variables are set the same way from `getColumnMinWidths()`/`getColumnMaxWidths()` (`Code.js`).

The template also carries a `mode` variable, set by the server function that opens the dialog (`template.mode = ...` before `.evaluate()`) and surfaced as `const INITIAL_MODE = <?!= JSON.stringify(mode); ?>;` (must be the unescaped `<?!=` form — `JSON.stringify()`'s quotes would otherwise be HTML-escaped and corrupt the token). `openWebEditor()` sets `mode = null` (normal list-view editor); `openPhotoExport()` sets `'photoExport'`, which `init()` checks first to skip the normal `getSchemaAndData()` bootstrap. **Every server function that calls `HtmlService.createTemplateFromFile('WebEditor')` must set `template.mode` explicitly** — an unset scriptlet variable throws `ReferenceError`.

### Client ↔ server communication

`google.script.run` is the only way the client talks to the server. Every call is asynchronous and must chain `.withSuccessHandler()` and `.withFailureHandler()` before the function name:

```js
google.script.run
  .withSuccessHandler(result => { ... })
  .withFailureHandler(err => { ... })
  .serverFunctionName(arg1, arg2);
```

The client code (the `WebEditor.*.js.html` fragments) is not a module — everything is a global within the `HtmlService` sandbox.

### Opening remote spreadsheets safely (`openSpreadsheetSafely`)

A `SpreadsheetApp.openById(id)` that throws because the user can't access `id` makes the *whole* enclosing execution fail (e.g. `getSchemaAndData()` returns "You do not have permission to access the required document") **even if you catch it** — an Apps Script runtime quirk. `openSpreadsheetSafely(id)` (`Code.js`) avoids it: it probes with `DriveApp.getFileById(id)` (caught, `null` on failure — the same pattern `getImagesDataUrls()` uses for `"no-access"` images) and only then calls `SpreadsheetApp.openById(id)`, returning `null` on any failure.

**Every** call site that opens a spreadsheet by an ID not guaranteed accessible to the current user must use it: `Handbook!B2:B` master sources, the `Handbook!A6` actual-personnel link, and client-supplied `spreadsheetId`s in `movePersonnel`/`updateRow`/`deleteRow`/`Export.js`.

Related helpers (`Utils.js`) — use them instead of repeating the inline pattern:

- `getDatabaseSheet(spreadsheetId)` → `{ ss, sheet }`; wraps `resolveSpreadsheet` + `getSheetByName(SHEET_DATABASE)` and throws on failure. All write functions needing the `Database` sheet (`addRowWithData`, `updateRow`, `deleteRow`, `deleteRows`, `movePersonnel` destination) use it; destructure only what you need.
- `getDriveIdFromHandbook(handbookSheet, cellAddress)` — trims the cell and pipes it through `parseDriveId()`; returns `''` if the sheet is missing. Use it for any Handbook cell holding a Drive ID or sharing URL.
- `findKeyByPattern(obj, pattern)` / `getFieldByPattern(data, pattern)` — the `COL_DRAFT_DATE`, `COL_SERVICE_HISTORY`, `COL_CLOSE_RELATIVES`, `COL_MARITAL_STATUS`, `COL_CONTRACT_UNTIL`, `COL_PHONE_NUMBER` (and `COL_FULL_NAME`/`COL_PHOTO`/`COL_CARD_ID`) constants in `Config.js` are case-insensitive **regexes** (e.g. `/дата призову/i`) so header matching tolerates casing drift across Master Mode sources. They can't be used as an object key (`data[COL_X]`) or compared with `===`; use `getFieldByPattern()` for the cell value, `findKeyByPattern()` when the literal matched header text is needed (as `_fillServiceHistoryTable()` does), and `COL_PHONE_NUMBER.test(...)` for plain matching (as `fixPhoneNumbers()` does).
- `parseDriveId()` (server, `Utils.js`) and `extractDriveId()` (client, `WebEditor.images.js.html`) share URL patterns — update both together; `tests/drive-id.test.js` checks they agree.

**`PERMISSION_DENIED` despite real access:** check for multiple signed-in Google accounts in the browser before chasing a code fix — the dialog's `google.script.run` calls can authenticate against the wrong account, giving the same error as a genuine sharing gap. Confirmed fix once: log out of all Google accounts except the one with access.

### Decision records

Deliberate-but-surprising choices are recorded in `docs/decisions/` so they aren't re-derived: [0001 open spreadsheets via a Drive probe](docs/decisions/0001-open-spreadsheets-via-drive-probe.md), [0002 lint the client fragments with an inline processor](docs/decisions/0002-lint-client-fragments-with-inline-processor.md), [0003 client script fragments share one global scope](docs/decisions/0003-client-script-fragments-single-global-scope.md).

### Feature docs (read the relevant one before changing that area)

- [docs/architecture-master-mode.md](docs/architecture-master-mode.md) — Master Mode loading (`getMasterSourceRows`, `queueMasterSourceFetch`, schema-mismatch warning), Refresh/Reset, source ordering, the Actual personnel filter, Move personnel, per-unit Drive folder resolution (`getUnitDataFolder`, `extractUnitName`).
- [docs/architecture-export-import.md](docs/architecture-export-import.md) — F-1/WC Docs export (4-pass order), XLSX export, S-КАДР photo export (two-phase, `mode = 'photoExport'`), award import.
- [docs/architecture-web-editor.md](docs/architecture-web-editor.md) — edit view tabs/panels, unsaved-changes and validation flow, `*-table` widths, dropdown/typed sub-column types, image loading/caching, Trash soft delete, `fixPhoneNumbers`/`fixFullNames`.

The user-facing reference (for people using/deploying the tool, not for changing its code) lives in `docs/` too, linked from the README: `spreadsheet-setup.md` (Database/Handbook layout, column types), `features.md`, `exports.md`, `configuration.md` (the `Config.js` constant table). When you add/rename a `Config.js` constant, a Handbook cell, a column type, a menu item, or a user-visible feature, update the matching one of these; keep the README itself short.

Rules that apply even if you don't open those docs:

- Client selection actions (export/move/delete) use `getSelectedRowEntries()`/`getSelectedRows()`; both exclude off-list rows when the Actual personnel filter is on.
- Move: hard-deleting source rows shifts every later `rowIndex` in that source; adjust the cached `schema.rows` (see `onDeleteSuccess()`/`onMoveSuccess()`).
- XLSX export: never `Blob.getAs()` for Sheet→xlsx, never a mixed text/formula `setFormulas()` call — see the doc.
- `_exportDoc()` pass order (images → service history → text columns → correspondence) is load-bearing.
- Never leave a persistent error element in `#edit-header`; use `alert()`.
- Don't destroy hidden `.edit-panel`s — `collectFormValues()` relies on them staying mounted.

## Constants — always in `Config.js`

All tuneable values belong in `Config.js`. Constants needed by the client must also be threaded through the `getSchemaAndData()` return value in `Code.js` (see `filterDebounceMs`, `imageFetchBatchSize`, `imageFetchConcurrency`, `masterMode` as examples). Never hardcode magic numbers in the `WebEditor.*.js.html` fragments.

**Sheet cell/column/row references in `Config.js` use A1 notation**, matching the Handbook range constants (`MASTER_MODE_CELL = 'A2'`, `HANDBOOK_TABLE_COLUMNS_RANGE = 'H2:J'`, `MASTER_MODE_SOURCES_RANGE = 'B2:B'`) — a 0-based numeric index isn't self-describing against the sheet. This includes single-column constants (`AWARDS_IMPORT_ID_COL = 'A'`, not `0`). When code needs a 0-based array index for a raw `getValues()` row, convert once via `columnLetterToIndex(letter)` (`Utils.js`).
