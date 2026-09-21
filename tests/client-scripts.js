'use strict';

// Single source of truth for the web editor's client script fragments: the
// include order in WebEditor.html is both the browser load order and the order
// ESLint concatenates them in. Shared by tests/load.js and eslint.config.mjs.

const fs = require('node:fs');
const path = require('node:path');

const SRC_DIR = path.join(__dirname, '..', 'src');
const SHELL_FILE = 'WebEditor.html';
const CLIENT_INCLUDE_REGEX = /createHtmlOutputFromFile\('([^']+\.js)'\)/g;

/**
 * Lists the client script fragments in load order, as WebEditor.html includes
 * them. Each fragment is a file under `srcDir` holding one <script> block.
 * @param {string} [srcDir] - Directory holding the fragments; defaults to src/.
 * @returns {Array<{file: string, text: string}>} `file` is the name under `srcDir`, `text` its full contents.
 */
function clientScripts(srcDir = SRC_DIR) {
  const shell = fs.readFileSync(path.join(srcDir, SHELL_FILE), 'utf8');
  return [...shell.matchAll(CLIENT_INCLUDE_REGEX)].map(([, name]) => {
    const file = `${name}.html`;
    return { file, text: fs.readFileSync(path.join(srcDir, file), 'utf8') };
  });
}

module.exports = { clientScripts, SRC_DIR };
