# 0001 — Open remote spreadsheets via a Drive probe, not a bare `openById`

## Context

`SpreadsheetApp.openById(id)` throws when the current user can't access `id`. In Apps Script, that failure makes the **whole enclosing execution** fail (for example `getSchemaAndData()` returns "You do not have permission to access the required document") even when the call sits inside a `try/catch`. Master Mode reads spreadsheets that some users legitimately can't open, and one inaccessible source must not take the editor down.

## Decision

Every call site that opens a spreadsheet by an ID not guaranteed accessible to the current user goes through `openSpreadsheetSafely(id)` (`Code.js`). It probes with `DriveApp.getFileById(id)` first (a failure there *is* catchable, and returns `null`), and only then calls `SpreadsheetApp.openById(id)`. Callers treat `null` as "no access".

The ESLint rule `no-restricted-properties` (in `eslint.config.mjs`) bans `SpreadsheetApp.openById`/`openByUrl` everywhere in `src/*.js`; `openSpreadsheetSafely()` carries the single `eslint-disable` comment.

## Consequences

- A raw `openById` looks simpler and is what the Apps Script docs show; the lint rule stops it being reintroduced "for cleanliness".
- One extra Drive call per open. Acceptable: opens are already cached per spreadsheet within an execution.
- `getImagesDataUrls()` uses the same probe pattern to report `"no-access"` images.
