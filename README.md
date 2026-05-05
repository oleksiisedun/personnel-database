# Personnel Database

A Google Sheets–based personnel database with a built-in web editor. Data lives in a Google Sheet; the web editor provides a richer UI for browsing and editing records.

## How it works

The project is a [Google Apps Script](https://developers.google.com/apps-script) bound to a Google Spreadsheet, deployed locally with [CLASP](https://github.com/google/clasp).

- **`Code.js`** — server-side script (menu, data access, image proxy, config constants)
- **`WebEditor.html`** — client app shell; includes CSS and JS via `<?!= HtmlService.createHtmlOutputFromFile(...) ?>`
- **`WebEditor.css.html`** — styles for the web editor
- **`WebEditor.js.html`** — client-side logic for the web editor

## Spreadsheet structure

### `Database` sheet

| Row | Purpose |
|-----|---------|
| 1 | Column names |
| 2 | Column types (`text`, `image`, `relatives-table`, `service-table`, …) |
| 3+ | Data rows |

### `Handbook` sheet

Defines sub-column headers for `*-table` column types.

| Column A | Column B onward |
|----------|----------------|
| `relatives-table` | Sub-column headers for that table type |
| `service-table` | Sub-column headers for that table type |

Row 1 is a header row (skipped). Each subsequent row maps a type name to its headers.

## Column types

| Type | List view | Edit view |
|------|-----------|-----------|
| `text` | Plain text | Text input |
| `image` | Thumbnail (click to enlarge) | Google Drive link input with live preview |
| `*-table` | Decoded mini-table | Row/column editor with add & delete |

### `*-table` encoding format

Table data is stored in a single cell as a pipe-and-newline delimited string:

```
value1 | value2 | value3
value1 | value2 | value3
```

### Image columns

Images are stored as Google Drive sharing links. Supported URL formats:

- `https://drive.google.com/file/d/FILE_ID/view`
- `https://drive.google.com/open?id=FILE_ID`
- Bare file ID

Images are fetched server-side (via `DriveApp`) and returned as base64 data URLs, so all users with access to the spreadsheet can view images regardless of their personal Drive session.

## Web editor features

- **List view** — full-screen table with all columns and data
- **Filtering** — debounced live filter input above each `text` column; supports plain text and regular expressions (toggle per column)
- **Empty-cell filter** — dropdown above each `*-table` column: All / Empty / Not empty
- **Add person** — button to append a new empty row and open it in the edit view immediately
- **Column visibility** — "Columns ▾" button to hide/show individual columns; first column is always visible
- **Image thumbnails** — loaded asynchronously, cached for the session
- **Lightbox** — click any thumbnail to view the full image
- **Edit view** — click a name in the first column to open a per-record editor

## Configuration

Both constants live at the top of `Code.js` and are injected into the client at render time.

### `EDIT_MODE`

```js
const EDIT_MODE = true;
```

Set to `false` to make the editor read-only: the "Add person" button is hidden and the first-column link that opens the edit view is disabled.

### `COLUMN_MIN_WIDTHS`

```js
const COLUMN_MIN_WIDTHS = { text: 150, image: 150, table: 900 };
```

Controls the minimum width (px) of each column type in the list view.

## Local development

```bash
# Install CLASP globally
npm install -g @google/clasp

# Authenticate
clasp login

# Push changes to the bound script project
clasp push

# Open the script editor in the browser
clasp open
```

The `.clasp.json` file already contains the script ID linking this directory to the deployed project.

### Deploying to multiple spreadsheets

Script IDs for all target spreadsheets are listed in `clasp-targets.json`:

```json
{
  "target1": "<script-id>",
  "target2": "<script-id>"
}
```

Run `clasp-push.sh` to push to all targets in sequence:

```bash
./clasp-push.sh
```

The script temporarily swaps the `scriptId` in `.clasp.json` for each target and restores the original on exit.
