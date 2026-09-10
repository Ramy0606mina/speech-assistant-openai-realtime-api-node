// A shared, text-only report model keeps email, PDF and Word output consistent.
// Source HTML is never executed or copied into the outgoing email.
export function cleanReportText(value) {
  return String(value ?? '').replace(/&#(x[0-9a-f]+|\d+);/gi, (all, code) => {
    const n = code[0].toLowerCase() === 'x' ? parseInt(code.slice(1), 16) : Number(code);
    return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : all;
  }).replace(/&nbsp;/gi, ' ').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"').replace(/&amp;/gi, '&')
    .replace(/\\([*#|`._-])/g, '$1').replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/`([^`]+)`/g, '$1').replace(/\u00a0/g, ' ');
}

export function reportSubject(value) {
  return String(value || 'London report').replace(/^(?:(?:re|fw|fwd)\s*:\s*|LONDON\s*[—–-]?\s*(?:Task (?:Complete|Response|Needs Attention)|Reminder Needs Attention)\s*[|:]?\s*)+/i, '').trim() || 'London report';
}

function cells(line) {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|')) s = s.slice(0, -1);
  return s.split('|').map(v => v.trim());
}
const separator = line => cells(line).every(v => /^:?-{3,}:?$/.test(v));

export function parseReport(text) {
  const lines = cleanReportText(text).split(/\r?\n/);
  const blocks = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    if (line.includes('|') && i + 1 < lines.length && separator(lines[i + 1])) {
      const columns = cells(line); const rows = [];
      i += 2;
      for (; i < lines.length && lines[i].trim().includes('|'); i++) {
        const row = cells(lines[i]);
        if (row.length !== columns.length) throw new Error('Report table has inconsistent columns; regenerate before delivery.');
        rows.push(row);
      }
      i--;
      if (columns.length < 2 || columns.length > 8) throw new Error('Report tables require 2 to 8 columns.');
      blocks.push({ type: 'table', columns, rows });
    } else if (line.startsWith('|') && line.split('|').length > 2) throw new Error('Report table is missing its header separator; regenerate before delivery.');
    else if (/^#{1,6}\s+/.test(line)) blocks.push({ type: 'heading', text: line.replace(/^#{1,6}\s+/, '') });
    else blocks.push({ type: 'paragraph', text: line });
  }
  return blocks;
}

export const escapeHtml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
export const isMoney = value => /^\(?-?\$[\d,]+(?:\.\d{2})?\)?$/.test(String(value).trim());

export function renderReportHtml(text) {
  const content = parseReport(text).map(block => {
    if (block.type === 'heading') return `<h2 style="font-size:17px;margin:22px 0 10px;color:#111827">${escapeHtml(block.text)}</h2>`;
    if (block.type === 'paragraph') return `<p style="margin:10px 0;line-height:1.5">${escapeHtml(block.text)}</p>`;
    const row = (values, header, index) => `<tr>${values.map(value => `<${header ? 'th' : 'td'} style="border:1px solid #cbd5e1;padding:9px 11px;vertical-align:top;text-align:${isMoney(value) ? 'right' : 'left'};background:${header ? '#e5eaf0' : index % 2 ? '#f5f7f9' : '#ffffff'}">${escapeHtml(value)}</${header ? 'th' : 'td'}>`).join('')}</tr>`;
    return `<table style="border-collapse:collapse;width:100%;font-size:14px;margin:12px 0"><thead>${row(block.columns, true, 0)}</thead><tbody>${block.rows.map((values, i) => row(values, false, i)).join('')}</tbody></table>`;
  }).join('');
  return `<!doctype html><html><body><div style="max-width:1000px;font-family:Arial,Helvetica,sans-serif;color:#111827;font-size:14px">${content}</div></body></html>`;
}
