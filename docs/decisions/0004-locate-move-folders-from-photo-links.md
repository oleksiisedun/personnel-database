# 0004 — Locate Move personnel folders from photo links

## Context

Move personnel also moves the person's Drive folder. It used to find the source unit folder by parsing the spreadsheet title (`"… - 7 РОП"` → `"7 РОП"`) and scanning the shared UNITS folder for that name, then scanning that for a subfolder named exactly like the person's full name. Any drift in a spreadsheet title, folder name or person name silently skipped the folder.

## Decision

Use the links already in the row instead of names:

- **Person folder:** the photo-column cell. A folder link is the person folder; a file link resolves to its single parent (`resolvePersonFolderId()`).
- **Destination folder:** the parent of the first resolvable person folder among the first `MOVE_DEST_FOLDER_PROBE_ROWS` rows of the destination Database (`findUnitFolderIdFromPhotos()`), read before the moved rows are appended.
- **No fallback** to name search. If anything is missing or ambiguous (no link, inaccessible, several parents, empty destination) the row still moves and the log says why the folder did not.

`Handbook!A4` (`DATA_FOLDER`), `UNIT_NAME_SEPARATOR`, `extractUnitName()` and `findFolderByNameCaseInsensitive()` were removed.

## Consequences

- Renaming people, spreadsheets or folders no longer breaks Move.
- A photo stored outside the person's own folder moves the wrong folder; a folder shared by several people moves them all. Nothing detects this.
- A destination with no photo links among its first rows cannot receive folders until one exists.
