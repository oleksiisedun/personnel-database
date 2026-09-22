'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer, plain } = require('./load.js');

const {
  normalizeFullName, normalizePhoneNumber, extractUnitName, columnLetterToIndex,
  compareColumnSchemas, extractColumnSchema, findColumnIndex, findKeyByPattern,
  getFieldByPattern, padRowToColumnCount, formatDateDDMMYYYY, groupAndSortBySpreadsheetId,
  parseGidFromUrl, UNIT_NAME_SEPARATOR, COL_FULL_NAME, COL_PHONE_NUMBER, COL_DRAFT_DATE,
} = loadServer(['Config.js', 'Formatting.js', 'DriveHelpers.js', 'SchemaHelpers.js', 'Utils.js'], [
  'normalizeFullName', 'normalizePhoneNumber', 'extractUnitName', 'columnLetterToIndex',
  'compareColumnSchemas', 'extractColumnSchema', 'findColumnIndex', 'findKeyByPattern',
  'getFieldByPattern', 'padRowToColumnCount', 'formatDateDDMMYYYY', 'groupAndSortBySpreadsheetId',
  'parseGidFromUrl', 'UNIT_NAME_SEPARATOR', 'COL_FULL_NAME', 'COL_PHONE_NUMBER', 'COL_DRAFT_DATE',
]);

describe('normalizeFullName', () => {
  test('uppercases the surname and keeps the rest as-is', () => {
    assert.equal(normalizeFullName('Шевченко Тарас Григорович'), 'ШЕВЧЕНКО Тарас Григорович');
  });
  test('trims and collapses spaces, tabs, newlines and non-breaking spaces', () => {
    assert.equal(normalizeFullName('  Шевченко \t Тарас\n Григорович  '), 'ШЕВЧЕНКО Тарас Григорович');
  });
  test('handles a single word and empty input', () => {
    assert.equal(normalizeFullName('шевченко'), 'ШЕВЧЕНКО');
    assert.equal(normalizeFullName('   '), '');
  });
  test('is idempotent', () => {
    const once = normalizeFullName(' іваненко   Іван ');
    assert.equal(normalizeFullName(once), once);
  });
});

describe('normalizePhoneNumber', () => {
  test('adds a leading zero to a bare 9-digit number', () => {
    assert.equal(normalizePhoneNumber('501234567'), '0501234567');
  });
  test('strips the 38 prefix from a 12-digit number', () => {
    assert.equal(normalizePhoneNumber('380501234567'), '0501234567');
  });
  test('leaves valid, malformed and empty values unchanged', () => {
    for (const phone of ['0501234567', '012345678', '38050123456', '3805012345678', '+380501234567', '050 123 45 67', '']) {
      assert.equal(normalizePhoneNumber(phone), phone);
    }
  });
});

describe('extractUnitName', () => {
  test('returns the text after the last separator, trimmed', () => {
    const sep = UNIT_NAME_SEPARATOR;
    assert.equal(extractUnitName(`Personnel${sep}Alpha${sep} 3rd Battalion `), '3rd Battalion');
  });
  test('falls back to the trimmed whole name when there is no separator', () => {
    assert.equal(extractUnitName('  Just a name '), 'Just a name');
  });
});

describe('columnLetterToIndex', () => {
  test('maps single and multi-letter columns to 0-based indices', () => {
    assert.equal(columnLetterToIndex('A'), 0);
    assert.equal(columnLetterToIndex('Z'), 25);
    assert.equal(columnLetterToIndex('AA'), 26);
    assert.equal(columnLetterToIndex('AL'), 37);
  });
  test('is case-insensitive', () => {
    assert.equal(columnLetterToIndex('ab'), columnLetterToIndex('AB'));
  });
});

describe('compareColumnSchemas', () => {
  const col = (name, type = 'text') => ({ name, type });

  test('returns nothing for identical schemas', () => {
    assert.deepEqual(plain(compareColumnSchemas([col('A'), col('B')], [col('A'), col('B')])), []);
  });
  test('reports a renamed column and a retyped column', () => {
    const result = plain(compareColumnSchemas([col('A'), col('B')], [col('A'), col('C', 'number')]));
    assert.deepEqual(result, [{ colIndex: 1, localName: 'B', localType: 'text', remoteName: 'C', remoteType: 'number' }]);
  });
  test('reports a column missing on one side', () => {
    const result = plain(compareColumnSchemas([col('A')], [col('A'), col('B')]));
    assert.deepEqual(result, [{ colIndex: 1, localName: '', localType: '', remoteName: 'B', remoteType: 'text' }]);
  });
  test('ignores trailing blank columns on both sides', () => {
    assert.deepEqual(plain(compareColumnSchemas([col('A'), col('', '')], [col('A')])), []);
  });
});

describe('extractColumnSchema', () => {
  test('pairs names with lowercased types', () => {
    assert.deepEqual(plain(extractColumnSchema([['ПІБ', 'Фото'], ['Text', 'IMAGE']])), [
      { name: 'ПІБ', type: 'text' }, { name: 'Фото', type: 'image' },
    ]);
  });
  test('gives an empty type when the type row is short or missing', () => {
    assert.deepEqual(plain(extractColumnSchema([['A', 'B'], ['text']])), [
      { name: 'A', type: 'text' }, { name: 'B', type: '' },
    ]);
    assert.deepEqual(plain(extractColumnSchema([['A']])), [{ name: 'A', type: '' }]);
  });
});

describe('header pattern lookup', () => {
  test('findColumnIndex matches case-insensitively on trimmed headers', () => {
    assert.equal(findColumnIndex(['№', ' Номер Телефону '], COL_PHONE_NUMBER), 1);
    assert.equal(findColumnIndex(['№', 'ПІБ'], COL_PHONE_NUMBER), -1);
  });
  test('findColumnIndex tolerates non-string header cells', () => {
    assert.equal(findColumnIndex([null, 42, 'ПІБ'], COL_FULL_NAME), 2);
  });
  test('findKeyByPattern returns the first matching key or null', () => {
    assert.equal(findKeyByPattern({ a: 1, 'Дата призову (ТЦК)': 2 }, COL_DRAFT_DATE), 'Дата призову (ТЦК)');
    assert.equal(findKeyByPattern({ a: 1 }, COL_DRAFT_DATE), null);
  });
  test('getFieldByPattern returns the value, or empty string when no key matches', () => {
    assert.equal(getFieldByPattern({ 'ДАТА ПРИЗОВУ': '01.02.2020' }, COL_DRAFT_DATE), '01.02.2020');
    assert.equal(getFieldByPattern({}, COL_DRAFT_DATE), '');
  });
});

describe('padRowToColumnCount', () => {
  test('pads short rows with empty strings and truncates long ones', () => {
    assert.deepEqual(plain(padRowToColumnCount(['a'], 3)), ['a', '', '']);
    assert.deepEqual(plain(padRowToColumnCount(['a', 'b', 'c'], 2)), ['a', 'b']);
    assert.deepEqual(plain(padRowToColumnCount([], 0)), []);
  });
  test('does not mutate the input', () => {
    const row = ['a'];
    padRowToColumnCount(row, 3);
    assert.deepEqual(row, ['a']);
  });
});

describe('formatDateDDMMYYYY', () => {
  test('zero-pads day and month', () => {
    assert.equal(formatDateDDMMYYYY(new Date(2024, 0, 5)), '05.01.2024');
  });
  test('does not pad two-digit parts', () => {
    assert.equal(formatDateDDMMYYYY(new Date(2024, 11, 31)), '31.12.2024');
  });
});

describe('parseGidFromUrl', () => {
  test('reads gid from the query string or the fragment', () => {
    assert.equal(parseGidFromUrl('https://docs.google.com/spreadsheets/d/X/edit?gid=123'), 123);
    assert.equal(parseGidFromUrl('https://docs.google.com/spreadsheets/d/X/edit#gid=456'), 456);
  });
  test('keeps gid=0 (first tab) distinct from a missing gid', () => {
    assert.equal(parseGidFromUrl('https://docs.google.com/spreadsheets/d/X/edit#gid=0'), 0);
    assert.equal(parseGidFromUrl('https://docs.google.com/spreadsheets/d/X/edit'), null);
  });
});

describe('groupAndSortBySpreadsheetId', () => {
  test('groups by spreadsheet and sorts each group by descending rowIndex', () => {
    const groups = groupAndSortBySpreadsheetId([
      { rowIndex: 3, spreadsheetId: 'a' },
      { rowIndex: 9, spreadsheetId: 'b' },
      { rowIndex: 7, spreadsheetId: 'a' },
      { rowIndex: 5, spreadsheetId: 'a' },
    ]);
    assert.deepEqual(plain([...groups.entries()]), [
      ['a', [{ rowIndex: 7, spreadsheetId: 'a' }, { rowIndex: 5, spreadsheetId: 'a' }, { rowIndex: 3, spreadsheetId: 'a' }]],
      ['b', [{ rowIndex: 9, spreadsheetId: 'b' }]],
    ]);
  });
  test('puts local rows (null or missing spreadsheetId) in one null group', () => {
    const groups = groupAndSortBySpreadsheetId([
      { rowIndex: 4, spreadsheetId: null },
      { rowIndex: 6 },
    ]);
    assert.deepEqual(plain([...groups.keys()]), [null]);
    assert.deepEqual(plain(groups.get(null).map(e => e.rowIndex)), [6, 4]);
  });
});
