# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Deployment

```bash
# Push to the primary bound script project
clasp push

# Push to all target spreadsheets listed in clasp-targets.json
./clasp-push.sh
```

There is no build step, linter, or test suite.

## Architecture

This is a **Google Apps Script** project (V8 runtime) bound to a Google Spreadsheet. The web editor runs inside an `HtmlService` modal dialog opened from the Sheets menu.

### How the HTML template works

`WebEditor.html` is served via `HtmlService.createTemplateFromFile()`. It includes CSS and JS using scriptlet tags:

```html
<?!= HtmlService.createHtmlOutputFromFile('WebEditor.css').getContent(); ?>
<?!= HtmlService.createHtmlOutputFromFile('WebEditor.js').getContent(); ?>
```

The `!` in `<?!=` means the content is included **without** HTML sanitization — the raw `<style>` and `<script>` tags are injected as-is. Column width CSS variables are also set via scriptlets in `WebEditor.html` using `getColumnMinWidths()` / `getColumnMaxWidths()` from `Code.js`.

### Client ↔ server communication

`google.script.run` is the only way the client talks to the server. Every call is asynchronous and must chain `.withSuccessHandler()` and `.withFailureHandler()` before the function name:

```js
google.script.run
  .withSuccessHandler(result => { ... })
  .withFailureHandler(err => { ... })
  .serverFunctionName(arg1, arg2);
```

The client-side code in `WebEditor.js.html` is not a module — all functions are globals within the `HtmlService` sandbox.

### Image loading

Images are fetched server-side via `DriveApp` (using the script owner's OAuth token) so users without personal Drive access can still view images. The client caches results in `imageCache` for the session.

Loading uses a **concurrency pool**: `IMAGE_FETCH_CONCURRENCY` batches of `IMAGE_FETCH_BATCH_SIZE` files run in parallel; when a batch completes the next one starts. Both constants live in `Config.js` and are passed to the client via `getSchemaAndData()`. The tradeoff:
- All at once → exceeds Apps Script's ~30 concurrent-execution limit (bottom images fail)
- Fully sequential → reliable but slow
- Pool of 3 × 10 → safe and fast

`DriveApp.getFileById()` throws for inaccessible files; the server catches it and returns `{ type: 'no-access' }`. The client renders a gray "No access" badge.

## Code conventions

### Constants — always in `Config.js`

All tuneable values belong in `Config.js`. Constants needed by the client must also be threaded through the `getSchemaAndData()` return value in `Code.js` (see `filterDebounceMs`, `imageFetchBatchSize`, `imageFetchConcurrency` as examples). Never hardcode magic numbers in `WebEditor.js.html`.
