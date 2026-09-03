import test from 'node:test';
import assert from 'node:assert/strict';

import { resolveFileContentType } from '../file-content-type.js';

test('replaces Dropbox generic MIME types using the filename extension', () => {
  const cases = [
    ['lease.pdf', 'application/pdf'],
    ['budget.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    ['contract.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    ['presentation.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
    ['photo.JPG', 'image/jpeg'],
  ];

  for (const [filename, expected] of cases) {
    assert.equal(
      resolveFileContentType({ filename, reportedContentType: 'application/octet-stream' }),
      expected
    );
  }
});

test('preserves a useful reported MIME type and removes parameters', () => {
  assert.equal(
    resolveFileContentType({
      filename: 'report.bin',
      reportedContentType: 'application/pdf; charset=binary',
    }),
    'application/pdf'
  );
});

test('uses a PDF signature when the filename has no extension', () => {
  assert.equal(
    resolveFileContentType({
      filename: 'scanned-contract',
      reportedContentType: 'application/octet-stream',
      buffer: Buffer.from('%PDF-1.7\\n'),
    }),
    'application/pdf'
  );
});

test('keeps the generic type only when the file cannot be identified', () => {
  assert.equal(
    resolveFileContentType({
      filename: 'unknown',
      reportedContentType: 'application/octet-stream',
      buffer: Buffer.from([0x00, 0x01, 0x02, 0x03]),
    }),
    'application/octet-stream'
  );
});
