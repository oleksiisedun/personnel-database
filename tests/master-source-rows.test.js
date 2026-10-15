'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer, plain } = require('./load.js');

const FILES = ['Config.js', 'Utils.js', 'SchemaHelpers.js', 'Formatting.js', 'Code.js'];

const HEADER = [
  ['Name', 'Age'],
  ['text', 'number'],
];

/**
 * Stand-in for a spreadsheet whose Database sheet holds `data` (or has no such sheet when null).
 * @param {Array<Array<*>>|null} data
 * @returns {object}
 */
const spreadsheetOf = (data) => ({
  getName: () => 'Unit A',
  getSheetByName: () => (data ? { getDataRange: () => ({ getValues: () => data }) } : null),
});

/**
 * Loads getMasterSourceRows() against a local sheet and a single remote one.
 * @param {{remote?: object|Error|null, remoteData?: Array<Array<*>>|null}} options - `remote`: what openById gives; null/Error makes DriveApp refuse.
 * @returns {Function}
 */
function loadWith({ remote, remoteData }) {
  const local = spreadsheetOf(HEADER);
  const { getMasterSourceRows } = loadServer(FILES, ['getMasterSourceRows'], {
    DriveApp: {
      getFileById: () => {
        if (remote === null) throw new Error('no access');
      },
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => local,
      openById: () => remote ?? spreadsheetOf(remoteData),
    },
  });
  return getMasterSourceRows;
}

describe('getMasterSourceRows loadError', () => {
  test("is 'no-access' when the spreadsheet can't be opened", () => {
    const r = plain(loadWith({ remote: null })('X'));
    assert.equal(r.loadError, 'no-access');
    assert.deepEqual(r.rows, []);
    assert.equal(r.name, 'X');
  });
  test("is 'no-database-sheet' when the Database sheet is missing", () => {
    assert.equal(loadWith({ remoteData: null })('X').loadError, 'no-database-sheet');
  });
  test('is null for a sheet with only header rows (empty, not an error)', () => {
    const r = loadWith({ remoteData: HEADER })('X');
    assert.equal(r.loadError, null);
    assert.equal(r.rows.length, 0);
  });
  test('is null and returns rows when the schema matches', () => {
    const r = plain(loadWith({ remoteData: [...HEADER, ['Ann', 30]] })('X'));
    assert.equal(r.loadError, null);
    assert.deepEqual(r.rows, [{ rowIndex: 3, values: ['Ann', '30'], spreadsheetId: 'X' }]);
  });
  test('is null (but columnMismatches set) when the schema differs', () => {
    const r = loadWith({
      remoteData: [
        ['Name', 'Rank'],
        ['text', 'text'],
      ],
    })('X');
    assert.equal(r.loadError, null);
    assert.equal(r.columnMismatches.length, 1);
  });
  test("is 'error' when reading the sheet throws", () => {
    const bad = {
      getName: () => 'U',
      getSheetByName: () => {
        throw new Error('boom');
      },
    };
    assert.equal(loadWith({ remote: bad })('X').loadError, 'error');
  });
});
