'use strict';

// Smoke test for the split client script: runs every WebEditor*.js.html fragment
// in include order inside one vm context with minimal browser stubs, the way the
// browser evaluates consecutive classic <script> tags (top-level let/const/function
// share one global scope). Catches a load-time reference to something declared in a
// later fragment (a TDZ/ReferenceError) and a fragment that fails to parse.
// It does not run any UI code; the dialog itself can only be checked by hand.

const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { clientScripts } = require('./load.js');

/**
 * Evaluates the fragments in the given order in a fresh context.
 * @param {Array<{file: string, text: string}>} fragments
 * @returns {{context: object, listeners: Array<[string, Function]>}}
 */
function runFragments(fragments) {
  const listeners = [];
  const context = vm.createContext({
    google: { script: { host: { setWidth() {}, setHeight() {} }, run: {} } },
    screen: { availWidth: 1200, availHeight: 800 },
    document: { addEventListener: (type, fn) => listeners.push([type, fn]) },
  });
  for (const { file, text } of fragments) {
    const body = text.slice(text.indexOf('<script>') + '<script>'.length, text.lastIndexOf('</script>'));
    vm.runInContext(body, context, { filename: file });
  }
  return { context, listeners };
}

test('all fragments load in include order without throwing', () => {
  assert.doesNotThrow(() => runFragments(clientScripts()));
});

test('registers init() for DOMContentLoaded and defines the entry points other code calls', () => {
  const { context, listeners } = runFragments(clientScripts());
  const [type, fn] = listeners[0];
  assert.equal(type, 'DOMContentLoaded');
  assert.equal(fn, vm.runInContext('init', context));
  for (const name of ['openEditView', 'renderList', 'runExport', 'runMove', 'queueImageFetch', 'buildMiniTable']) {
    assert.equal(vm.runInContext(`typeof ${name}`, context), 'function', `${name}() is not defined`);
  }
});
