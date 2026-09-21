'use strict';

// Calendar-accurate service duration used by the F-1 total-service-length value.

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer, plain } = require('./load.js');

const { _calendarDuration, _computeTotalServiceLength } = loadServer(
  ['Config.js', 'Utils.js', 'ExportValues.js'], ['_calendarDuration', '_computeTotalServiceLength']);

const ymd = (y, m, d) => new Date(y, m - 1, d);

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
