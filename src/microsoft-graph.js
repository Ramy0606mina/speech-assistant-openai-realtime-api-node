import { fetchJson } from './http.js';

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

export class MicrosoftGraphClient {
  constructor({
    readTenantId,
    readClientId,
    readClientSecret,
    actionTenantId,
    actionClientId,
    actionClientSecret,
    londonMailbox,
    ramyMailbox,
    fetchImpl = fetch,
  }) {
    this.readCreds = { tenantId: readTenantId, clientId: readClientId, clientSecret: readClientSecret };
    this.actionCreds = {
      tenantId: actionTenantId || readTenantId,
      clientId: actionClientId || readClientId,
      clientSecret: actionClientSecret || readClientSecret,
    };
    this.londonMailbox = londonMailbox;
    this.ramyMailbox = ramyMailbox;
    this.fetchImpl = fetchImpl;
    this.readToken = { value: '', expiresAt: 0 };
    this.actionToken = { value: '', expiresAt: 0 };
  }

  async #getToken(creds, cache) {
    if (cache.value && cache.expiresAt > Date.now() + 60000) return cache.value;
    if (!creds.tenantId || !creds.clientId || !creds.clientSecret) {
      throw new Error('Microsoft Graph credentials are not configured.');
    }
    const form = new URLSearchParams({
      client_id: creds.clientId,
      client_secret: creds.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    });
    const data = await fetchJson(
      this.fetchImpl,
      `https://login.microsoftonline.com/${encodeURIComponent(creds.tenantId)}/oauth2/v2.0/token`,
      { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: form.toString() },
      12000,
    );
    cache.value = data.access_token;
    cache.expiresAt = Date.now() + Math.max(60, Number(data.expires_in || 3600) - 120) * 1000;
    return cache.value;
  }

  get readMailbox() { return normalizeEmail(this.londonMailbox); }
  get principalMailbox() { return normalizeEmail(this.ramyMailbox); }

  async listLondonInbox(limit = 10) {
    if (!this.londonMailbox) throw new Error('LONDON_MINACO_EMAIL is not configured.');
    const token = await this.#getToken(this.readCreds, this.readToken);
    const url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.londonMailbox)}/mailFolders/inbox/messages`);
    url.searchParams.set('$top', String(Math.min(25, Math.max(1, Number(limit) || 10))));
    url.searchParams.set('$select', 'id,internetMessageId,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments');
    url.searchParams.set('$orderby', 'receivedDateTime desc');
    const data = await fetchJson(this.fetchImpl, url, { headers: { Authorization: `Bearer ${token}` } });
    return data?.value || [];
  }

  async getLondonMessage(messageId) {
    if (!messageId) throw new Error('messageId is required.');
    const token = await this.#getToken(this.readCreds, this.readToken);
    const url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.londonMailbox)}/messages/${encodeURIComponent(messageId)}`);
    url.searchParams.set('$select', 'id,internetMessageId,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,hasAttachments');
    return fetchJson(this.fetchImpl, url, {
      headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.body-content-type="text"' },
    });
  }

  async sendMail({ to, subject, body, cc = [] }) {
    if (!to) throw new Error('Recipient is required.');
    const token = await this.#getToken(this.actionCreds, this.actionToken);
    const recipients = (Array.isArray(to) ? to : [to]).filter(Boolean).map((address) => ({ emailAddress: { address } }));
    const ccRecipients = (Array.isArray(cc) ? cc : [cc]).filter(Boolean).map((address) => ({ emailAddress: { address } }));
    await fetchJson(this.fetchImpl, `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.londonMailbox)}/sendMail`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: {
          subject: String(subject || '').trim() || '(no subject)',
          body: { contentType: 'Text', content: String(body || '') },
          toRecipients: recipients,
          ccRecipients,
        },
        saveToSentItems: true,
      }),
    });
    return { sent: true };
  }
}
