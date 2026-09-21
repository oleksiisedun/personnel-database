'use strict';

// Client-side bookkeeping for cached sheet row numbers after rows are deleted
// or moved out of a source sheet (see onDeleteSuccess()/onMoveSuccess()).

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadClientFunction, plain } = require('./load.js');

const groupRowIndexesBySource = loadClientFunction('groupRowIndexesBySource');
const adjustRowIndexesAfterDelete = loadClientFunction('adjustRowIndexesAfterDelete');

const row = (rowIndex, spreadsheetId) => (spreadsheetId === undefined ? { rowIndex } : { rowIndex, spreadsheetId });

describe('groupRowIndexesBySource', () => {
  test('groups indexes per spreadsheet and treats missing/undefined id as the local (null) source', () => {
    const grouped = groupRowIndexesBySource([row(5), row(7, null), row(3, 'A'), row(9, 'A')]);
    assert.deepEqual(plain([...grouped]), [[null, [5, 7]], ['A', [3, 9]]]);
  });
});

describe('adjustRowIndexesAfterDelete', () => {
  test('shifts only rows below a deleted row, once per deleted row above them', () => {
    // Sheet rows 4 and 6 were deleted; the cache holds only the surviving rows.
    const rows = [row(3), row(5), row(7), row(10)];
    adjustRowIndexesAfterDelete(rows, new Map([[null, [4, 6]]]));
    assert.deepEqual(rows.map(r => r.rowIndex), [3, 4, 5, 8]);
  });
  test('leaves other spreadsheets untouched', () => {
    const rows = [row(5, 'A'), row(5, 'B'), row(5)];
    adjustRowIndexesAfterDelete(rows, new Map([['A', [3]]]));
    assert.deepEqual(rows.map(r => r.rowIndex), [4, 5, 5]);
  });
  test('treats undefined and null spreadsheetId as the same local source', () => {
    const rows = [row(8), row(9, null)];
    adjustRowIndexesAfterDelete(rows, new Map([[null, [3]]]));
    assert.deepEqual(rows.map(r => r.rowIndex), [7, 8]);
  });
  test('does nothing for an empty deletion map', () => {
    const rows = [row(3), row(4)];
    adjustRowIndexesAfterDelete(rows, new Map());
    assert.deepEqual(rows.map(r => r.rowIndex), [3, 4]);
  });
});
