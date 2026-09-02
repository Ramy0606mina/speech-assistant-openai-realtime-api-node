const MIME_BY_EXTENSION = Object.freeze({
  pdf: 'application/pdf',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  rtf: 'application/rtf',
  odt: 'application/vnd.oasis.opendocument.text',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  odp: 'application/vnd.oasis.opendocument.presentation',
  csv: 'text/csv',
  tsv: 'text/tab-separated-values',
  txt: 'text/plain',
  md: 'text/markdown',
  json: 'application/json',
  xml: 'application/xml',
  html: 'text/html',
  htm: 'text/html',
  yaml: 'application/yaml',
  yml: 'application/yaml',
  log: 'text/plain',
  js: 'text/javascript',
  mjs: 'text/javascript',
  cjs: 'text/javascript',
  ts: 'text/typescript',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  gif: 'image/gif',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  bmp: 'image/bmp',
  heic: 'image/heic',
  heif: 'image/heif',
  svg: 'image/svg+xml',
  zip: 'application/zip',
});

const GENERIC_CONTENT_TYPES = new Set([
  '',
  'application/octet-stream',
  'binary/octet-stream',
  'application/binary',
  'application/force-download',
]);

const normalizeContentType = (value) =>
  String(value || '')
    .split(';', 1)[0]
    .trim()
    .toLowerCase();

const extensionFromFilename = (filename) => {
  const basename = String(filename || '').split(/[\\\\/]/).at(-1) || '';
  const lastDot = basename.lastIndexOf('.');
  return lastDot > -1 ? basename.slice(lastDot + 1).toLowerCase() : '';
};

const detectImageOrPdfSignature = (buffer) => {
  if (!Buffer.isBuffer(buffer) || buffer.length < 4) return '';

  if (buffer.subarray(0, 5).toString('ascii') === '%PDF-') return 'application/pdf';
  if (buffer.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    return 'image/png';
  }
  if (buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) return 'image/jpeg';
  if (buffer.subarray(0, 4).toString('ascii') === 'GIF8') return 'image/gif';
  if (
    buffer.length >= 12 &&
    buffer.subarray(0, 4).toString('ascii') === 'RIFF' &&
    buffer.subarray(8, 12).toString('ascii') === 'WEBP'
  ) {
    return 'image/webp';
  }

  return '';
};

export const resolveFileContentType = ({ filename, reportedContentType, buffer } = {}) => {
  const reported = normalizeContentType(reportedContentType);
  if (!GENERIC_CONTENT_TYPES.has(reported)) return reported;

  const extensionType = MIME_BY_EXTENSION[extensionFromFilename(filename)];
  if (extensionType) return extensionType;

  return detectImageOrPdfSignature(buffer) || 'application/octet-stream';
};
