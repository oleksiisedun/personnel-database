# Personnel Database

A Google Sheets–based personnel database with a built-in web editor. Data lives in a Google Sheet; the web editor provides a richer UI for browsing, editing, and exporting records.

To set it up, see [Getting started](#getting-started). Detailed reference (spreadsheet layout, features, exports, configuration) is under [Documentation](#documentation).

## How it works

The project is a [Google Apps Script](https://developers.google.com/apps-script) bound to a Google Spreadsheet, deployed locally with [CLASP](https://github.com/google/clasp). The Sheets menu (`onOpen()` in `Code.js`) opens the web editor as a modal dialog; the browser-side UI (`WebEditor.*.html`) talks to the server only through `google.script.run`. `Code.js` handles data access and row CRUD, delegates document, XLSX and photo exports to `Export.js`, and shares helpers from `Utils.js` with `Export.js` and `Import.js`. All data lives in the `Database`, `Handbook` and `Trash` sheets (plus remote spreadsheets in Master Mode) and in Google Drive.

```mermaid
flowchart TD
    subgraph client["Browser (client)"]
        html["WebEditor.html<br/>App shell"]
        js["WebEditor.js.html<br/>Filter, edit, export UI"]
        css["WebEditor.css.html<br/>Styles"]
        html --> js
        html --> css
    end

    subgraph server["Server — Google Apps Script"]
        config["Config.js<br/>Constants & IDs"]
        code["Code.js<br/>Menu, data access, image proxy, openSpreadsheetSafely"]
        utils["Utils.js<br/>Shared helpers"]
        export["Export.js<br/>F-1, WC, XLSX & photo exports"]
        import_["Import.js<br/>Award import from S-КАДР"]
    end

    subgraph workspace["Google Workspace"]
        subgraph sheets["Google Sheets (data store)"]
            database["Database<br/>Personnel rows + column types"]
            handbook["Handbook<br/>Schema, dropdowns, Master Mode config"]
            trash["Trash<br/>Soft-deleted records"]
        end
        drive["Google Drive<br/>Photos, F-1/WC templates, exported docs"]
    end

    remote["Remote spreadsheets<br/>(Master Mode sources)"]

    subgraph devtools["Local dev tooling"]
        clasp["CLASP<br/>Push / pull to GAS"]
        claspPush["clasp-push.sh<br/>Multi-spreadsheet deploy"]
        manifest["appsscript.json<br/>Manifest & OAuth scopes"]
    end

    js -- "google.script.run" --> code
    code -. return data .-> js
    code --> config
    code --> database
    code --> handbook
    code --> trash
    code --> drive
    code -. "Master Mode read" .-> remote
    code -- "batch export" --> export
    code --> utils
    export --> utils
    import_ --> utils
    export --> drive
    import_ -. "Master Mode read/write" .-> remote
    import_ --> database

    clasp -.->|push/pull| code
    claspPush --> clasp
    manifest -.-> code

    classDef code   fill:#EEEDFE,stroke:#534AB7,color:#26215C
    classDef client fill:#E1F5EE,stroke:#0F6E56,color:#04342C
    classDef store  fill:#FFF8E1,stroke:#F9A825,color:#3E2723
    classDef export fill:#FAECE7,stroke:#993C1D,color:#4A1B0C
    classDef devtools fill:#F1EFE8,stroke:#5F5E5A,color:#2C2C2A

    class html,js,css client
    class config,code,utils code
    class database,handbook,trash,drive,remote store
    class export,import_ export
    class clasp,claspPush,manifest devtools
```

All deployable code lives in `src/` (the only directory clasp pushes); tooling and docs stay at the repo root.

| File | Purpose |
|------|---------|
| `src/Config.js` | All constants — sheet names, column names, Drive IDs, export settings |
| `src/Code.js` | Server-side script: menu, data access, image proxy |
| `src/Utils.js` | Shared server-side helpers (spreadsheet resolution, schema comparison, column lookups) used by `Code.js`, `Export.js`, and `Import.js` |
| `src/Export.js` | Server-side export logic for F-1 and Wanted Card documents, the single-file XLSX export, and the S-КАДР photo export |
| `src/Import.js` | Server-side award import from an external S-КАДР sheet |
| `src/appsscript.json` | Apps Script manifest — time zone, V8 runtime, and the Advanced Drive Service used for export thumbnails |
| `src/WebEditor.html` | Client app shell; includes CSS and JS via `<?!= HtmlService.createHtmlOutputFromFile(...) ?>` |
| `src/WebEditor.css.html` | Styles for the web editor |
| `src/WebEditor.js.html` | Client-side logic for the web editor |
| `docs/` | User-facing reference (spreadsheet setup, features, exports, configuration) and contributor-level `architecture-*.md` notes per feature area |
| `eslint.config.mjs`, `jsconfig.json` | Lint and typecheck config for `npm run check` |
| `test/` | Node unit tests (`npm test`) and the `vm` loader that runs the Apps Script sources under Node; not deployed |
| `clasp-push.sh`, `clasp-targets.json` | Multi-spreadsheet deploy script and its (git-ignored) target list |

## Features

- **List view** with per-column filters (plain text or regex), column visibility, image thumbnails, and an optional "Actual personnel" filter
- **Edit view** per person, with typed fields (dates, TIN, numbers, dropdowns, sub-tables), save-time validation, and a warning before unsaved edits are discarded
- **Add / delete** — new rows go to the local `Database` sheet; deleted rows move to `Trash` rather than being destroyed
- **Document export** — fill F-1 and Wanted Card Google Docs templates from the selected rows, or export them as a single `.xlsx`
- **Master Mode** — aggregate several unit spreadsheets in one editor and move personnel between them
- **Menu utilities** — fix phone numbers and full names in place; in Master Mode, S-КАДР photo export and award import

See [Features](docs/features.md) for the details.

## Getting started

### Prerequisites

- A Google account that can edit the target spreadsheet, which needs `Database` and `Handbook` sheets (see [Spreadsheet structure](docs/spreadsheet-setup.md#spreadsheet-structure) and [Sample files](docs/spreadsheet-setup.md#sample-files))
- [Node.js](https://nodejs.org) 20.19+ / 22.13+ (required by ESLint 10) and npm
- [`jq`](https://jqlang.github.io/jq/) — only needed for `clasp-push.sh`

### Install and connect

```bash
npm install                    # local tooling: TypeScript + ESLint
npm install -g @google/clasp   # the Apps Script CLI
clasp login
```

Create `.clasp.json` in the repo root (it's git-ignored). Use the script ID from the spreadsheet's **Extensions → Apps Script → Project Settings**:

```json
{ "scriptId": "<script-id>", "rootDir": "src" }
```

`rootDir` is required: it makes clasp push only `src/` (including `src/appsscript.json`) and never the repo-root tooling files. Re-add it if you ever regenerate the file with `clasp clone`/`clasp create`.

```bash
clasp push          # deploy src/ to the bound script project
clasp open-script   # open the Apps Script editor in the browser
```

`src/appsscript.json` (the manifest) is pushed along with the code and enables the [Advanced Drive Service](https://developers.google.com/apps-script/advanced/drive), which document export uses to fetch resized photo thumbnails — there's nothing to switch on by hand.

Then reload the spreadsheet: the **More... ⭐️** menu appears, and **Open Web Editor** launches the editor.

### First run

1. In the spreadsheet, choose **More... ⭐️ → Open Web Editor**. Google asks you to authorize the script the first time.
2. The list view shows every row of the `Database` sheet. Type in a filter box above a column to narrow the list (plain text, or regex if you toggle it).
3. Click a name in the first column to open that person's edit view, change a field and press **Save**; **← Back** returns to the list.
4. Press **Add Person** to append a new empty row and edit it straight away. **Delete** in the edit view moves a record to the `Trash` sheet.
5. Tick the checkboxes at the left of some rows, then **Export XLSX** to save them as one `.xlsx` in the Drive export folder. **Export F-1** / **Export Wanted Card** fill the Docs templates instead. All exports need the export-folder cell in `Handbook` filled in, and the document exports also need their template cells (see [Exports](docs/exports.md)).

### Deploying to multiple spreadsheets

Script IDs for all target spreadsheets are listed in `clasp-targets.json` (git-ignored — create it yourself on a fresh clone):

```json
{
  "target1": "<script-id>",
  "target2": "<script-id>"
}
```

Run `clasp-push.sh` (or `npm run clasp-push`) to `clasp push --force` to every target in sequence:

```bash
./clasp-push.sh
```

The script temporarily swaps the `scriptId` in `.clasp.json` for each target and restores the original on exit.

## Development

```bash
npm run check       # typecheck + lint, in sequence
npm run typecheck   # tsc over src/*.js against @types/google-apps-script (non-strict)
npm run lint        # ESLint over src/*.js and the <script> in WebEditor.js.html
npm test            # node --test unit tests for the pure helpers (no dependencies)
```

There is no build step. Unit tests in `test/` cover the pure helpers (`Utils.js`) and keep the server and client Drive-ID parsers in sync; they run under Node, not Apps Script. Because JSDoc is the only source of type info, a typecheck failure often means a stale `@param`/`@returns`. Lint's main job is `no-undef` on `WebEditor.js.html`, which `tsc` can't see into. Feature-level internals are in [`docs/`](docs/).

## Documentation

- [Spreadsheet setup](docs/spreadsheet-setup.md) — `Database`, `Trash` and `Handbook` layout, column types, sample files
- [Features](docs/features.md) — Master Mode, custom menu, S-КАДР photo export and award import, web editor features
- [Exports](docs/exports.md) — F-1 / Wanted Card documents, computed values, XLSX export
- [Configuration](docs/configuration.md) — every `Config.js` constant
- Contributors: [Master Mode](docs/architecture-master-mode.md), [export and import](docs/architecture-export-import.md) and [web editor](docs/architecture-web-editor.md) internals
