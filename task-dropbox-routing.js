const DEFAULT_DROPBOX_ROOT = '/LONDON - ACCESS';

const SUPPORTED_EXTENSION_PATTERN =
  'pdf|docx?|xlsx?|xlsm|csv|tsv|pptx?|rtf|odt|ods|odp|txt|md|json|xml|html?|ya?ml|log|jpe?g|png|gif|webp|tiff?|bmp|heic|heif|svg';

const SUPPORTED_FILE_PATTERN = new RegExp(
  `\\.(${SUPPORTED_EXTENSION_PATTERN})$`,
  'i'
);

const FILE_REFERENCE_PATTERN = new RegExp(
  `\\/LONDON\\s*-\\s*ACCESS\\/[^\\r\\n<>\"]+?\\.(?:${SUPPORTED_EXTENSION_PATTERN})\\b`,
  'gi'
);

const normalizeInstructionText = (value) => {
  let text = String(value || '')
    .replace(/&#x20;|&#32;|&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/\\_/g, '_');

  for (let index = 0; index < 4; index += 1) {
    const joined = text.replace(
      /(\/LONDON\s*-\s*ACCESS[^\r\n]*)\r?\n(?=[^\r\n/]+\/)/gi,
      '$1 '
    );
    if (joined === text) break;
    text = joined;
  }

  return text;
};

const trimReference = (value) =>
  String(value || '')
    .trim()
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[),;:.]+$/g, '')
    .replace(/\\+$/g, '');

const canonicalPath = (value) =>
  trimReference(value)
    .replace(/\\/g, '/')
    .replace(/\/{2,}/g, '/')
    .toLowerCase();

const parentPath = (value) => {
  const path = trimReference(value).replace(/\\/g, '/').replace(/\/{2,}/g, '/');
  const lastSlash = path.lastIndexOf('/');
  return lastSlash > 0 ? path.slice(0, lastSlash) : path;
};

const uniquePaths = (values) => {
  const seen = new Set();
  const result = [];
  for (const value of values) {
    const cleaned = trimReference(value);
    const key = canonicalPath(cleaned);
    if (!cleaned || seen.has(key)) continue;
    seen.add(key);
    result.push(cleaned);
  }
  return result;
};

export const isSupportedDropboxAnalysisFile = (filename) =>
  SUPPORTED_FILE_PATTERN.test(String(filename || '').trim());

export const extractDropboxReferencePlan = (
  instruction,
  dropboxRoot = DEFAULT_DROPBOX_ROOT
) => {
  const text = normalizeInstructionText(instruction);
  const rootPattern = new RegExp(
    String(dropboxRoot || DEFAULT_DROPBOX_ROOT)
      .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
      .replace(/\\ /g, '\\s*'),
    'i'
  );

  const explicitFiles = uniquePaths(text.match(FILE_REFERENCE_PATTERN) || []);
  const folders = explicitFiles.map(parentPath);

  for (const line of text.split(/\r?\n/)) {
    const rootMatch = rootPattern.exec(line);
    if (!rootMatch) continue;

    const candidate = trimReference(line.slice(rootMatch.index));
    if (!candidate || isSupportedDropboxAnalysisFile(candidate)) continue;
    if (/\/(?:LONDON\s*-\s*ACCESS)(?:\/|$)/i.test(candidate)) {
      folders.push(candidate.replace(/\/$/, ''));
    }
  }

  return {
    referenced: explicitFiles.length > 0 || folders.length > 0,
    explicitFiles,
    folders: uniquePaths(folders),
  };
};

const extensionOf = (filename) =>
  String(filename || '').split('.').at(-1)?.toLowerCase() || '';

const instructionTokens = (instruction) =>
  new Set(
    normalizeInstructionText(instruction)
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter((token) => token.length >= 4)
      .filter(
        (token) =>
          ![
            'access',
            'dropbox',
            'folder',
            'files',
            'file',
            'same',
            'with',
            'from',
            'that',
            'this',
            'into',
          ].includes(token)
      )
  );

export const selectDropboxTaskFiles = ({
  entries,
  explicitFiles,
  instruction,
  maxFiles = 10,
} = {}) => {
  const explicit = new Set((explicitFiles || []).map(canonicalPath));
  const text = normalizeInstructionText(instruction).toLowerCase();
  const tokens = instructionTokens(instruction);
  const deduped = new Map();

  for (const entry of Array.isArray(entries) ? entries : []) {
    if (entry?.type && entry.type !== 'file') continue;
    const path = trimReference(entry?.path || '');
    const name = String(entry?.name || path.split('/').at(-1) || '').trim();
    if (!path || !name || !isSupportedDropboxAnalysisFile(name)) continue;

    const key = canonicalPath(path);
    if (!deduped.has(key)) deduped.set(key, { ...entry, path, name });
  }

  for (const path of explicitFiles || []) {
    const cleaned = trimReference(path);
    const key = canonicalPath(cleaned);
    if (!deduped.has(key)) {
      deduped.set(key, {
        type: 'file',
        path: cleaned,
        name: cleaned.split('/').at(-1) || cleaned,
        size: 0,
      });
    }
  }

  const scored = [...deduped.values()].map((entry) => {
    const key = canonicalPath(entry.path);
    const nameLower = entry.name.toLowerCase();
    const extension = extensionOf(entry.name);
    let score = explicit.has(key) ? 10000 : 0;

    if (text.includes(nameLower)) score += 4000;
    if (extension === 'pdf' && /\bpdfs?\b/i.test(text)) score += 2500;
    if (['doc', 'docx'].includes(extension) && /\b(docx?|word|report)\b/i.test(text)) {
      score += 1800;
    }
    if (['xls', 'xlsx', 'xlsm', 'csv', 'tsv'].includes(extension) && /\b(excel|spreadsheet|workbook|csv)\b/i.test(text)) {
      score += 1800;
    }
    if (['ppt', 'pptx'].includes(extension) && /\b(powerpoint|presentation|slides?)\b/i.test(text)) {
      score += 1800;
    }

    const nameTokens = nameLower.split(/[^a-z0-9]+/).filter((token) => token.length >= 4);
    score += nameTokens.filter((token) => tokens.has(token)).length * 150;
    return { ...entry, score };
  });

  const relevant = scored.filter((entry) => entry.score > 0);
  return relevant
    .sort(
      (a, b) =>
        b.score - a.score ||
        Number(a.size || 0) - Number(b.size || 0) ||
        a.path.localeCompare(b.path)
    )
    .slice(0, Math.max(0, Number(maxFiles) || 0));
};
