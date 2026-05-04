# Personnel Database

A Google Sheets–based personnel database with a built-in web editor. Data lives in a Google Sheet; the web editor provides a richer UI for browsing and editing records.

## How it works

The project is a [Google Apps Script](https://developers.google.com/apps-script) bound to a Google Spreadsheet, deployed locally with [CLASP](https://github.com/google/clasp).

- **`Code.js`** — server-side script (menu, data access, image proxy)
- **`WebEditor.html`** — single-file client app (list view + edit view)

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
- **Filtering** — live filter input above each `text` column
- **Column visibility** — "Columns ▾" button to hide/show individual columns; first column is always visible
- **Image thumbnails** — loaded asynchronously, cached for the session
- **Lightbox** — click any thumbnail to view the full image
- **Edit view** — click a name in the first column to open a per-record editor

## Width constraints

Adjust these CSS variables at the top of `WebEditor.html`:

```css
:root {
  --col-min-width: 150px;        /* minimum width for text columns */
  --col-image-min-width: 150px;  /* minimum width for image columns */
  --col-table-min-width: 900px;  /* minimum width for *-table columns */
}
```

## Local development

```bash
# Install CLASP globally
npm install -g @google/clasp

# Authenticate
clasp login

# Push changes to the script project
clasp push

# Open the script editor in the browser
clasp open
```

The `.clasp.json` file already contains the script ID linking this directory to the deployed project.
