'use strict';

// parseDriveId() (server, DriveHelpers.js) and extractDriveId() (client, WebEditor.images.js.html)
// are hand-maintained copies of the same URL patterns; these tests keep them aligned.

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer, loadClientFunction } = require('./load.js');

const { parseDriveId, looksLikeDriveUrl } = loadServer(['Config.js', 'DriveHelpers.js'], ['parseDriveId', 'looksLikeDriveUrl']);
const extractDriveId = loadClientFunction('extractDriveId');

const ID = '1AbC_dEf-GhIjKlMnOpQrStUvWxYz012345';

// Every URL form the two docs promise to support, plus a few near-misses.
const URLS = [
  `https://drive.google.com/drive/folders/${ID}`,
  `https://drive.google.com/drive/u/0/folders/${ID}`,
  `https://drive.google.com/drive/folders/${ID}?usp=sharing`,
  `https://drive.google.com/file/d/${ID}/view`,
  `https://drive.google.com/file/d/${ID}/view?usp=drive_link`,
  `https://drive.google.com/open?id=${ID}`,
  `https://drive.google.com/uc?export=download&id=${ID}`,
  `https://docs.google.com/spreadsheets/d/${ID}/edit?gid=0#gid=0`,
];

describe('Drive ID extraction — server and client agree', () => {
  for (const url of URLS) {
    test(`both extract the ID from ${url}`, () => {
      assert.equal(parseDriveId(url), ID);
      assert.equal(extractDriveId(url), ID);
    });
  }

  test('both pass a bare ID through', () => {
    assert.equal(parseDriveId(ID), ID);
    assert.equal(extractDriveId(ID), ID);
  });

  test('both agree when a URL carries several ID-like segments', () => {
    const url = `https://drive.google.com/drive/folders/${ID}?id=OTHER_ID_12345`;
    assert.equal(parseDriveId(url), extractDriveId(url));
  });
});

describe('parseDriveId', () => {
  test('returns non-URL input unchanged, including surrounding whitespace', () => {
    assert.equal(parseDriveId(''), '');
    assert.equal(parseDriveId('not a url'), 'not a url');
  });
});

describe('looksLikeDriveUrl', () => {
  test('is true for URLs and false for bare IDs and plain text', () => {
    assert.equal(looksLikeDriveUrl(URLS[0]), true);
    assert.equal(looksLikeDriveUrl(ID), false);
    assert.equal(looksLikeDriveUrl('hello'), false);
  });
});

describe('extractDriveId', () => {
  test('returns null for empty or unparseable input', () => {
    assert.equal(extractDriveId(''), null);
    assert.equal(extractDriveId(null), null);
    assert.equal(extractDriveId('short'), null);
    assert.equal(extractDriveId('not a url at all'), null);
  });
  test('trims a bare ID', () => {
    assert.equal(extractDriveId(`  ${ID} `), ID);
  });
});
