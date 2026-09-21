'use strict';

// Pure helpers behind the F-1 / Wanted Card / XLSX exports: computed values,
// sub-table encoding, pluralization and escaping. None touch Apps Script services.

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer, plain } = require('./load.js');

const {
  _pluralizeUk, _calendarDuration, _computeTotalServiceLength, _computeContractSignDate,
  _computeValue, _computeChildrenNamesBirthDates, _parseSubTable, _encodeSubTable,
  _escapeRegex, _escapeReplacement, _escapeFormulaString, _colIndexToA1Column,
} = loadServer(['Config.js', 'Utils.js', 'ExportValues.js', 'ExportXlsx.js', 'Export.js'], [
  '_pluralizeUk', '_calendarDuration', '_computeTotalServiceLength', '_computeContractSignDate',
  '_computeValue', '_computeChildrenNamesBirthDates', '_parseSubTable', '_encodeSubTable',
  '_escapeRegex', '_escapeReplacement', '_escapeFormulaString', '_colIndexToA1Column',
]);

const ymd = (y, m, d) => new Date(y, m - 1, d);

describe('_pluralizeUk', () => {
  test('years', () => {
    const cases = { 0: 'років', 1: 'рік', 2: 'роки', 4: 'роки', 5: 'років', 11: 'років', 12: 'років', 14: 'років', 21: 'рік', 22: 'роки', 111: 'років' };
    for (const [n, word] of Object.entries(cases)) assert.equal(_pluralizeUk(Number(n), 'year'), `${n} ${word}`);
  });
  test('months', () => {
    const cases = { 0: 'місяців', 1: 'місяць', 3: 'місяці', 5: 'місяців', 11: 'місяців', 12: 'місяців', 21: 'місяць' };
    for (const [n, word] of Object.entries(cases)) assert.equal(_pluralizeUk(Number(n), 'month'), `${n} ${word}`);
  });
  test('days', () => {
    const cases = { 0: 'днів', 1: 'день', 3: 'дні', 11: 'днів', 14: 'днів', 22: 'дні', 25: 'днів', 31: 'день' };
    for (const [n, word] of Object.entries(cases)) assert.equal(_pluralizeUk(Number(n), 'day'), `${n} ${word}`);
  });
});

describe('_calendarDuration', () => {
  test('is zero on the same day and when end is before start', () => {
    assert.deepEqual(plain(_calendarDuration(ymd(2026, 5, 9), ymd(2026, 5, 9))), { years: 0, months: 0, days: 0 });
    assert.deepEqual(plain(_calendarDuration(ymd(2026, 5, 9), ymd(2026, 5, 8))), { years: 0, months: 0, days: 0 });
  });
  test('counts exact anniversaries as whole years', () => {
    assert.deepEqual(plain(_calendarDuration(ymd(2020, 5, 9), ymd(2026, 5, 9))), { years: 6, months: 0, days: 0 });
  });
  test('ordinary case borrows from the previous month', () => {
    assert.deepEqual(plain(_calendarDuration(ymd(2020, 1, 15), ymd(2026, 3, 10))), { years: 6, months: 1, days: 23 });
  });
  test('crosses a year boundary', () => {
    assert.deepEqual(plain(_calendarDuration(ymd(2025, 12, 20), ymd(2026, 1, 5))), { years: 0, months: 0, days: 16 });
  });
  test('regression: start day past the end of the previous month never gives negative days', () => {
    // 31 Jan -> 1 Mar used to yield { months: 1, days: -2 }.
    assert.deepEqual(plain(_calendarDuration(ymd(2026, 1, 31), ymd(2026, 3, 1))), { years: 0, months: 1, days: 1 });
    assert.deepEqual(plain(_calendarDuration(ymd(2026, 1, 30), ymd(2026, 3, 1))), { years: 0, months: 1, days: 1 });
    assert.deepEqual(plain(_calendarDuration(ymd(2024, 1, 31), ymd(2024, 3, 1))), { years: 0, months: 1, days: 1 });
  });
  test('month-end start that has not reached the next month-end yet', () => {
    assert.deepEqual(plain(_calendarDuration(ymd(2026, 1, 31), ymd(2026, 2, 28))), { years: 0, months: 0, days: 28 });
  });
  test('leap-day start', () => {
    assert.deepEqual(plain(_calendarDuration(ymd(2020, 2, 29), ymd(2021, 3, 1))), { years: 1, months: 0, days: 1 });
  });
  test('never returns out-of-range parts for any start/end pair over two years', () => {
    for (let s = 0; s < 731; s++) {
      const start = new Date(2023, 0, 1 + s);
      for (let e = 0; e < 800; e += 7) {
        const { years, months, days } = _calendarDuration(start, new Date(start.getFullYear(), start.getMonth(), start.getDate() + e));
        assert.ok(years >= 0 && months >= 0 && months <= 11 && days >= 0 && days <= 30,
          `${start.toDateString()} + ${e}d -> ${years}y ${months}m ${days}d`);
      }
    }
  });
});

describe('_computeTotalServiceLength', () => {
  test('formats the duration from the last date in the cell, as of the given day', () => {
    const data = { 'Дата призову': '01.01.2010; 15.01.2020' };
    assert.equal(_computeTotalServiceLength(data, ymd(2026, 3, 10)), '6 років, 1 місяць, 23 дні (станом на 10.03.2026)');
  });
  test('does not print negative days for a month-end start', () => {
    const data = { 'Дата призову': '31.01.2026' };
    assert.equal(_computeTotalServiceLength(data, ymd(2026, 3, 1)), '0 років, 1 місяць, 1 день (станом на 01.03.2026)');
  });
  test('returns an empty string when there is no date', () => {
    assert.equal(_computeTotalServiceLength({ 'Дата призову': '' }, ymd(2026, 3, 1)), '');
    assert.equal(_computeTotalServiceLength({}, ymd(2026, 3, 1)), '');
  });
});

describe('_computeContractSignDate', () => {
  const history = '01.01.2015 - 01.01.2020 | командир в/ч 3011 | ...';
  test('joins the first draft date with the unit number from the first service record', () => {
    const data = { 'Дата призову': '07.05.2015', 'Проходження служби': history };
    assert.equal(_computeContractSignDate(data), '07.05.2015 з в/ч 3011');
  });
  test('falls back to the default unit number without service history', () => {
    assert.equal(_computeContractSignDate({ 'Дата призову': '07.05.2015' }), '07.05.2015 з в/ч 3102');
  });
  test('is empty for mobilised personnel and when there is no date', () => {
    assert.equal(_computeContractSignDate({ 'Дата призову': '07.05.2015', 'Контракт укладено до': 'Мобілізований' }), '');
    assert.equal(_computeContractSignDate({ 'Дата призову': '07.05.2015', 'Контракт укладено до': 'мобілізована' }), '');
    assert.equal(_computeContractSignDate({ 'Проходження служби': history }), '');
  });
});

describe('_computeValue and relatives lookups', () => {
  // Built with the real encoder so empty fields are encoded exactly as the editor writes them.
  const data = {
    'Близькі родичі': _encodeSubTable([
      ['мати', 'Іванова Марія', 'м. Київ', '0501111111', '01.01.1960'],
      ['Чоловік', 'Петренко Іван', 'м. Львів', '0502222222', '02.02.1985'],
      ['дитина', 'Петренко Олег', 'м. Львів', '', '03.03.2015'],
      ['дитина', 'Петренко Анна', 'м. Львів', '', ''],
    ]),
  };
  test('dispatches by key and matches relation case-insensitively', () => {
    assert.equal(_computeValue('motherFullName', data), 'Іванова Марія');
    assert.equal(_computeValue('motherPhoneNumber', data), '0501111111');
    assert.equal(_computeValue('spouseFullName', data), 'Петренко Іван');
    assert.equal(_computeValue('fatherFullName', data), '');
  });
  test('unknown key returns an empty string', () => {
    assert.equal(_computeValue('nope', data), '');
  });
  test('numbers children and omits a missing birth date', () => {
    assert.equal(_computeChildrenNamesBirthDates(data), '1 дитина: Петренко Олег 03.03.2015\n2 дитина: Петренко Анна');
  });
});

describe('_parseSubTable / _encodeSubTable', () => {
  test('parses rows and fields', () => {
    assert.deepEqual(plain(_parseSubTable('a | b\nc | d')), [['a', 'b'], ['c', 'd']]);
  });
  test('returns [] for empty input and drops rows with a blank first field', () => {
    assert.deepEqual(plain(_parseSubTable('')), []);
    assert.deepEqual(plain(_parseSubTable(undefined)), []);
    assert.deepEqual(plain(_parseSubTable(' | x\nA | y\n')), [['A', 'y']]);
  });
  test('encode is the inverse of parse for well-formed input', () => {
    const cell = 'a | b | c\nd | e | f';
    assert.equal(_encodeSubTable(_parseSubTable(cell)), cell);
  });
});

describe('escaping helpers', () => {
  test('_escapeRegex makes every metacharacter literal', () => {
    const tricky = '[1+1] (a.b)* ^$ {x} | \\';
    assert.ok(new RegExp(_escapeRegex(tricky)).test(tricky));
    assert.equal(_escapeRegex('a.b'), 'a\\.b');
  });
  test('_escapeReplacement neutralizes backslash and dollar', () => {
    assert.equal(_escapeReplacement('$1 and \\n'), '\\$1 and \\\\n');
    assert.equal(_escapeReplacement('plain'), 'plain');
  });
  test('_escapeFormulaString doubles quotes', () => {
    assert.equal(_escapeFormulaString('say "hi"'), 'say ""hi""');
  });
});

describe('_colIndexToA1Column', () => {
  test('converts 0-based indices to column letters', () => {
    const cases = { 0: 'A', 25: 'Z', 26: 'AA', 51: 'AZ', 52: 'BA', 701: 'ZZ', 702: 'AAA' };
    for (const [index, letters] of Object.entries(cases)) assert.equal(_colIndexToA1Column(Number(index)), letters);
  });
});
