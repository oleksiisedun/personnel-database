'use strict';

// Pure helpers behind the award import from S-КАДР: grouping import-sheet rows
// by personnel ID and building the post-import summary text. Neither touches
// Apps Script services — the sheet/UI side (importAwards(), _applyAwards())
// is exercised manually, not here (see docs/unit-tests.md).

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer, plain } = require('./load.js');

const {
  _groupAwardsById, _buildImportSummary, AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT,
  // _groupAwardsById does `instanceof Date` on a cell value, so a Date built with
  // Node's own Date constructor (a different realm than the vm context) would
  // never match — the sheet's own Date constructor makes an instance it recognizes.
  Date: ContextDate,
} = loadServer(['Config.js', 'Formatting.js', 'SchemaHelpers.js', 'Utils.js', 'Import.js'], [
  '_groupAwardsById', '_buildImportSummary', 'AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT', 'Date',
]);

// Columns per Config.js: A=id, F=name, G=order number, H=order date. Row 1 is the header.
const row = (id, name, orderNumber, orderDate) => [id, '', '', '', '', name, orderNumber, orderDate];

describe('_groupAwardsById', () => {
  test('groups multiple award rows under the same ID', () => {
    const data = [
      ['header'],
      row('123', 'орден', '45/2024', '01.02.2024'),
      row('123', 'медаль', '46/2024', '02.02.2024'),
    ];
    assert.deepEqual(plain(_groupAwardsById(data)), {
      123: [
        ['Орден', '45', '01.02.2024'],
        ['Медаль', '46', '02.02.2024'],
      ],
    });
  });

  test('skips rows with no ID or no award name', () => {
    const data = [
      ['header'],
      row('', 'орден', '1', ''),
      row('123', '', '1', ''),
      row('123', 'орден', '1', ''),
    ];
    assert.deepEqual(plain(_groupAwardsById(data)), { 123: [['Орден', '1', '']] });
  });

  test('capitalizes only the first letter of the award name', () => {
    const data = [['header'], row('1', 'бойова медаль', '', '')];
    assert.equal(_groupAwardsById(data)['1'][0][0], 'Бойова медаль');
  });

  test('cleans the order number, stripping a trailing "/year" and other punctuation', () => {
    const data = [['header'], row('1', 'x', '№ 45/2024', '')];
    assert.equal(_groupAwardsById(data)['1'][0][1], '45');
  });

  test('formats a real Date cell as DD.MM.YYYY, passes a string cell through trimmed', () => {
    const data = [
      ['header'],
      row('1', 'x', '', new ContextDate(2024, 1, 3)),
      row('2', 'x', '', '  04.02.2024  '),
    ];
    const grouped = _groupAwardsById(data);
    assert.equal(grouped['1'][0][2], '03.02.2024');
    assert.equal(grouped['2'][0][2], '04.02.2024');
  });

  test('returns an empty object for a header-only sheet', () => {
    assert.deepEqual(plain(_groupAwardsById([['header']])), {});
  });
});

describe('_buildImportSummary', () => {
  const base = { peopleUpdated: 2, entriesAdded: 3, entriesDuplicate: 1, notFound: [] };

  test('omits the "Not found" section when nothing is missing', () => {
    const summary = _buildImportSummary(base);
    assert.match(summary, /People updated: 2/);
    assert.match(summary, /Award entries added: 3/);
    assert.match(summary, /Duplicate entries skipped: 1/);
    assert.doesNotMatch(summary, /Not found:/);
  });

  test('lists every missing ID, uncapped, with no "more" suffix', () => {
    const notFound = ['a', 'b', 'c'];
    const summary = _buildImportSummary({ ...base, notFound });
    assert.match(summary, /IDs not found: 3/);
    assert.match(summary, /Not found: a, b, c$/m);
    assert.doesNotMatch(summary, /more\)/);
  });

  test('caps the displayed list at AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT and appends a "+N more" suffix', () => {
    const notFound = Array.from({ length: AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT + 5 }, (_, i) => `id${i}`);
    const summary = _buildImportSummary({ ...base, notFound });
    assert.match(summary, new RegExp(`IDs not found: ${notFound.length}`));
    assert.match(summary, /\(\+5 more\)/);
    const shown = summary.match(/Not found: (.+?) \(\+5 more\)/)[1].split(', ');
    assert.equal(shown.length, AWARDS_IMPORT_NOT_FOUND_DISPLAY_LIMIT);
  });
});
