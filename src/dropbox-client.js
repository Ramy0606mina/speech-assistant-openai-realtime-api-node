import { fetchJson } from './http.js';

function normalizeDropboxPath(value) {
  let path = String(value || '').trim().replace(/\\/g, '/');
  if (!path.startsWith('/')) path = `/${path}`;
  path = path.replace(/\/{2,}/g, '/');
  return path.length > 1 ? path.replace(/\/$/, '') : path;
}

export class DropboxClient {
  constructor({ accessToken, rootPath = '/LONDON - ACCESS', fetchImpl = fetch }) {
    this.accessToken = accessToken;
    this.rootPath = normalizeDropboxPath(rootPath);
    this.fetchImpl = fetchImpl;
  }

  resolvePath(relativeOrAbsolute = '') {
    const candidate = normalizeDropboxPath(relativeOrAbsolute || this.rootPath);
    const rootLower = this.rootPath.toLowerCase();
    const candidateLower = candidate.toLowerCase();
    if (candidateLower === rootLower || candidateLower.startsWith(`${rootLower}/`)) return candidate;
    if (candidate === '/') return this.rootPath;
    const joined = normalizeDropboxPath(`${this.rootPath}/${String(relativeOrAbsolute || '').replace(/^\/+/, '')}`);
    if (!joined.toLowerCase().startsWith(`${rootLower}/`) && joined.toLowerCase() !== rootLower) {
      throw new Error('Dropbox path escapes configured London root.');
    }
    return joined;
  }

  async #rpc(endpoint, body) {
    if (!this.accessToken) throw new Error('DROPBOX_ACCESS_TOKEN is not configured.');
    return fetchJson(this.fetchImpl, `https://api.dropboxapi.com/2/${endpoint}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${this.accessToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
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
}

export { normalizeDropboxPath };
