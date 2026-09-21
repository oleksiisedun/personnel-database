'use strict';

// Static checks for contracts between the client files and the server that no
// type checker or linter can see, and that otherwise only fail at runtime in the
// modal dialog: element ids, toolbar button styling, and template variables.

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const read = name => fs.readFileSync(path.join(__dirname, '..', 'src', name), 'utf8');
const html = read('WebEditor.html');
const clientJs = read('WebEditor.js.html');
const codeJs = read('Code.js');

describe('WebEditor element ids', () => {
  const staticIds = new Set([...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
  // Ids assigned in script, e.g. `checkbox.id = 'chk-select-all'`.
  const dynamicIds = new Set([...clientJs.matchAll(/\.id\s*=\s*'([^']+)'/g)].map(m => m[1]));
  const looked = [...new Set([...clientJs.matchAll(/getElementById\('([^']+)'\)/g)].map(m => m[1]))];

  test('finds ids to check (guards against the regexes silently matching nothing)', () => {
    assert.ok(staticIds.size > 30 && looked.length > 30);
  });
  test('every literal getElementById() id exists in WebEditor.html or is assigned in script', () => {
    const missing = looked.filter(id => !staticIds.has(id) && !dynamicIds.has(id));
    assert.deepEqual(missing, []);
  });
  test('ids in WebEditor.html are unique', () => {
    const all = [...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]);
    const dupes = all.filter((id, i) => all.indexOf(id) !== i);
    assert.deepEqual(dupes, []);
  });
});

describe('toolbar buttons', () => {
  test('every <button> inside #toolbar has the btn-toolbar class (sizing/padding)', () => {
    const start = html.indexOf('id="toolbar"');
    const end = html.indexOf('<!-- Schema warning overlay -->');
    assert.ok(start !== -1 && end > start, 'toolbar region not found');
    const buttons = [...html.slice(start, end).matchAll(/<button\b[^>]*>/g)].map(m => m[0]);
    assert.ok(buttons.length > 5);
    const without = buttons.filter(tag => !/class="[^"]*\bbtn-toolbar\b/.test(tag));
    assert.deepEqual(without, []);
  });
});

describe('server functions that serve WebEditor', () => {
  test('each sets template.mode before evaluating (an unset scriptlet variable throws ReferenceError)', () => {
    const bodies = codeJs.split(/^function /m).filter(b => b.includes("createTemplateFromFile('WebEditor')"));
    assert.ok(bodies.length >= 2, 'expected openWebEditor() and openPhotoExport()');
    for (const body of bodies) {
      const name = body.slice(0, body.indexOf('('));
      assert.match(body, /\btemplate\.mode\s*=/, `${name}() does not set template.mode`);
    }
  });
});
