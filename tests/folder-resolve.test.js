'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadServer } = require('./load.js');

const { resolvePersonFolderId, findUnitFolderIdFromPhotos } = loadServer(
  ['Config.js', 'DriveHelpers.js'],
  ['resolvePersonFolderId', 'findUnitFolderIdFromPhotos'],
);

/** Builds a lookup over a plain map of id → {isFolder, parentIds}. */
const lookupIn = (items) => (id) => items[id] ?? null;

const ITEMS = {
  personFolder: { isFolder: true, parentIds: ['unitFolder'] },
  photo: { isFolder: false, parentIds: ['personFolder'] },
  orphanPhoto: { isFolder: false, parentIds: [] },
  sharedPhoto: { isFolder: false, parentIds: ['a', 'b'] },
  rootFolder: { isFolder: true, parentIds: [] },
};
const lookup = lookupIn(ITEMS);

describe('resolvePersonFolderId', () => {
  test('a folder link is the person folder itself', () => {
    assert.deepEqual({ ...resolvePersonFolderId('personFolder', lookup) }, { id: 'personFolder', reason: '' });
  });
  test('a file link resolves to its single parent', () => {
    assert.equal(resolvePersonFolderId('photo', lookup).id, 'personFolder');
  });
  test('accepts full sharing URLs', () => {
    assert.equal(resolvePersonFolderId('https://drive.google.com/file/d/photo/view?usp=sharing', lookup).id, 'personFolder');
    assert.equal(resolvePersonFolderId('https://drive.google.com/drive/folders/personFolder', lookup).id, 'personFolder');
  });
  test('reports why nothing was resolved', () => {
    assert.equal(resolvePersonFolderId('', lookup).reason, 'No photo link');
    assert.equal(resolvePersonFolderId('  ', lookup).reason, 'No photo link');
    assert.equal(resolvePersonFolderId('missing', lookup).reason, 'Photo not accessible');
    assert.equal(resolvePersonFolderId('orphanPhoto', lookup).reason, 'Photo has no single parent folder');
    assert.equal(resolvePersonFolderId('sharedPhoto', lookup).reason, 'Photo has no single parent folder');
  });
  test('never returns an id together with a reason', () => {
    for (const cell of ['photo', 'missing', '', 'sharedPhoto']) {
      const r = resolvePersonFolderId(cell, lookup);
      assert.ok(Boolean(r.id) !== Boolean(r.reason));
    }
  });
});

describe('findUnitFolderIdFromPhotos', () => {
  test('returns the parent of the first resolvable person folder', () => {
    assert.equal(findUnitFolderIdFromPhotos(['', 'missing', 'photo', 'personFolder'], lookup), 'unitFolder');
  });
  test('skips person folders without a single parent', () => {
    assert.equal(findUnitFolderIdFromPhotos(['rootFolder', 'photo'], lookup), 'unitFolder');
  });
  test('returns empty string when nothing resolves', () => {
    assert.equal(findUnitFolderIdFromPhotos([], lookup), '');
    assert.equal(findUnitFolderIdFromPhotos(['', 'missing', 'rootFolder'], lookup), '');
  });
});
