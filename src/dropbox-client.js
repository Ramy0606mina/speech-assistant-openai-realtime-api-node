import { fetchJson } from './http.js';
import { createHash } from 'node:crypto';
import mammoth from 'mammoth';
import { renderReportPdf } from './report-pdf.js';

function normalizeDropboxPath(value) {
  let path = String(value || '').trim().replace(/\\/g, '/');
  if (!path.startsWith('/')) path = `/${path}`;
  path = path.replace(/\/{2,}/g, '/');
  return path.length > 1 ? path.replace(/\/$/, '') : path;
}

export class DropboxClient {
  constructor({ accessToken, refreshToken, appKey, appSecret, saveReports = false, rootPath = '/LONDON - ACCESS', fetchImpl = fetch }) {
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
    this.appKey = appKey;
    this.appSecret = appSecret;
    this.tokenExpiresAt = 0;
    this.refreshInFlight = null;
    this.saveReports = saveReports;
    this.rootPath = normalizeDropboxPath(rootPath);
    this.fetchImpl = fetchImpl;
  }

  resolvePath(relativeOrAbsolute = '') {
    const candidate = normalizeDropboxPath(relativeOrAbsolute || this.rootPath);
    if (candidate.split('/').some(part => part === '..' || part === '.') || /[\u0000-\u001f]/.test(candidate)) {
      throw new Error('Dropbox path escapes configured London root.');
    }
    const rootLower = this.rootPath.toLowerCase();
    const candidateLower = candidate.toLowerCase();
    if (candidateLower === rootLower || candidateLower.startsWith(`${rootLower}/`)) return candidate;
    if (candidate === '/') return this.rootPath;
    if (String(relativeOrAbsolute).startsWith('/')) throw new Error('Dropbox path is outside configured London root.');
    const joined = normalizeDropboxPath(`${this.rootPath}/${String(relativeOrAbsolute || '').replace(/^\/+/, '')}`);
    if (!joined.toLowerCase().startsWith(`${rootLower}/`) && joined.toLowerCase() !== rootLower) {
      throw new Error('Dropbox path escapes configured London root.');
    }
    return joined;
  }

  async #token(force = false) {
    if (this.refreshToken && this.appKey) {
      if (!force && this.accessToken && this.tokenExpiresAt > Date.now() + 60000) return this.accessToken;
      if (!this.refreshInFlight) this.refreshInFlight = (async () => {
        const form = new URLSearchParams({ grant_type: 'refresh_token', refresh_token: this.refreshToken, client_id: this.appKey });
        if (this.appSecret) form.set('client_secret', this.appSecret);
        let data;
        try {
          data = await fetchJson(this.fetchImpl, 'https://api.dropboxapi.com/oauth2/token', {
            method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: form.toString(),
          });
        } catch { throw new Error('Dropbox credential renewal failed; reconnect the existing Dropbox app.'); }
        if (!data?.access_token) throw new Error('Dropbox renewal returned no access token.');
        this.accessToken = data.access_token;
        this.tokenExpiresAt = Date.now() + Number(data.expires_in || 14400) * 1000;
        return this.accessToken;
      })().finally(() => { this.refreshInFlight = null; });
      return this.refreshInFlight;
    }
    if (!this.accessToken) throw new Error('Dropbox runtime credentials are missing.');
    return this.accessToken;
  }

  async #rpc(endpoint, body, retried = false) {
    const token = await this.#token();
    try { return await fetchJson(this.fetchImpl, `https://api.dropboxapi.com/2/${endpoint}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    }); } catch (error) {
      if (error.status === 401 && !retried && this.refreshToken && this.appKey) {
        await this.#token(true);
        return this.#rpc(endpoint, body, true);
      }
      throw error;
    }
  }

  async listFolder(path = '') {
    const resolved = this.resolvePath(path);
    const data = await this.#rpc('files/list_folder', { path: resolved, recursive: false, include_deleted: false });
    return data?.entries || [];
  }

  async search(query, path = '') {
    const resolved = this.resolvePath(path);
    const data = await this.#rpc('files/search_v2', {
      query: String(query || '').trim(),
      options: { path: resolved, max_results: 50, filename_only: false },
    });
    return data?.matches || [];
  }

  async readFile(path, maxBytes = 40 * 1024 * 1024) {
    const resolved = this.resolvePath(path);
    const filename = resolved.split('/').at(-1);
    if (!/\.(pdf|docx?|xlsx?|pptx?|txt|csv|md|rtf)$/i.test(filename)) throw new Error('Unsupported Dropbox document type.');
    const token = await this.#token();
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);
    try {
      const response = await this.fetchImpl('https://content.dropboxapi.com/2/files/download', {
        method: 'POST', signal: controller.signal,
        headers: { Authorization: `Bearer ${token}`, 'Dropbox-API-Arg': JSON.stringify({ path: resolved }).replace(/[\u007f-\uffff]/g, c => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`) },
      });
      if (!response.ok) throw new Error(`Dropbox download failed (HTTP ${response.status}).`);
      const chunks = [];
      let size = 0;
      for await (const chunk of response.body) {
        size += chunk.length;
        if (size > maxBytes) { controller.abort(); throw new Error('Dropbox documents exceed the 40 MB task limit.'); }
        chunks.push(Buffer.from(chunk));
      }
      const bytes = Buffer.concat(chunks);
      if (!bytes.length) throw new Error('Dropbox returned an empty document.');
      if (filename.toLowerCase().endsWith('.docx')) {
        const extracted = await mammoth.extractRawText({ buffer: bytes });
        if (!extracted.value.trim()) throw new Error('Word document contains no extractable text; images and layout need separate review.');
        if (extracted.value.length > 500000) throw new Error('Word document exceeds the text analysis limit.');
        return { path: resolved, filename, size, part: { type: 'input_text',
          text: `Dropbox Word source: ${resolved}\nExtracted text only; images and visual layout are not included.\n\n${extracted.value}` } };
      }
      return { path: resolved, filename, size, part: {
        type: 'input_file', filename,
        file_data: `data:${filename.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'application/octet-stream'};base64,${bytes.toString('base64')}`,
      } };
    } finally { clearTimeout(timer); }
  }

  async saveReport({ taskKey, subject, text }) {
    if (!this.saveReports) throw new Error('Dropbox report saving is not enabled.');
    if (!taskKey || !String(text || '').trim()) throw new Error('Report task and content are required.');
    const folder = this.resolvePath('London Work');
    try {
      await this.#rpc('files/create_folder_v2', { path: folder, autorename: false });
    } catch (error) {
      if (error.status !== 409 || !String(error.data?.error_summary || '').startsWith('path/conflict/folder')) throw error;
    }
    // Content-addressed output names preserve prior reports and source documents.
    const contents = `# ${String(subject || 'London report').replace(/[\r\n]+/g, ' ')}\n\n${text}\n`;
    const suffix = createHash('sha256').update(`${taskKey}\n${contents}`).digest('hex');
    const label = String(subject || 'Report').replace(/[^a-zA-Z0-9 -]/g, '').trim().slice(0, 60) || 'Report';
    const path = this.resolvePath(`${folder}/London - ${label} - ${suffix.slice(0, 16)}.pdf`);
    const pdf = await renderReportPdf({ subject, text });
    const token = await this.#token();
    const data = await fetchJson(this.fetchImpl, 'https://content.dropboxapi.com/2/files/upload', {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/octet-stream',
        'Dropbox-API-Arg': JSON.stringify({ path, mode: 'overwrite', autorename: false, mute: true }).replace(/[\u007f-\uffff]/g, c => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`) },
      body: pdf,
    }, 30000);
    if (!data?.id || !data?.path_display) throw new Error('Dropbox did not confirm the saved report.');
    return { id: data.id, path: data.path_display };
  }
}

export { normalizeDropboxPath };
