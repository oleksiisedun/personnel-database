'use strict';

// Loads Apps Script sources into a Node vm context without modifying them.
// Server files share one global scope and have no modules, so they are run in
// order inside a single context; top-level `const`s live in that context's
// global lexical scope, hence evaluating an expression to read them back.

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const { clientScripts, SRC_DIR } = require('./client-scripts.js');

/**
 * Runs the given src/ files in one fresh context and returns the requested
 * top-level names (functions or constants) as an object.
 * @param {string[]} files - Filenames under src/, in load order (Config.js first).
 * @param {string[]} names - Top-level identifiers to return.
 * @returns {Object<string, *>}
 */
function loadServer(files, names) {
  const context = vm.createContext({});
  for (const file of files) {
    vm.runInContext(fs.readFileSync(path.join(SRC_DIR, file), 'utf8'), context, { filename: file });
  }
  return vm.runInContext(`({ ${names.join(', ')} })`, context);
}

/**
 * Returns a single function declared inside one of the client <script>
 * fragments. The scripts run DOM/`google.script` code at load time, so a
 * function can't be had by executing a whole fragment; instead its source is
 * cut out by its 2-space indent (declaration line to the next line that is
 * exactly `  }`) and evaluated alone. Only works for self-contained functions.
 * @param {string} name
 * @returns {Function}
 */
function loadClientFunction(name) {
  const pattern = new RegExp(`^  function ${name}\\(.*?^  }$`, 'ms');
  for (const { text } of clientScripts()) {
    const match = text.match(pattern);
    if (match) return vm.runInContext(`(${match[0].trim()})`, vm.createContext({}));
  }
  throw new Error(`Function ${name}() not found in any client script fragment`);
}

/**
 * Copies a value out of a vm context into this realm. Objects and arrays
 * created inside the context have a different prototype, which makes
 * assert.deepStrictEqual fail on otherwise-equal values.
 * @param {*} value - JSON-serializable value.
 * @returns {*}
 */
const plain = (value) => JSON.parse(JSON.stringify(value));

module.exports = { loadServer, loadClientFunction, clientScripts, plain };
