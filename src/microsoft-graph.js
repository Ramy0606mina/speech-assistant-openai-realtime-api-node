import { fetchJson } from './http.js';

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function clampLimit(value, fallback = 10, max = 25) {
  return Math.min(max, Math.max(1, Number(value) || fallback));
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

  async #listInbox(mailbox, limit = 10) {
    if (!mailbox) throw new Error('Mailbox is not configured.');
    const token = await this.#getToken(this.readCreds, this.readToken);
    const url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailbox)}/mailFolders/inbox/messages`);
    url.searchParams.set('$top', String(clampLimit(limit)));
    url.searchParams.set('$select', 'id,internetMessageId,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments');
    url.searchParams.set('$orderby', 'receivedDateTime desc');
    const data = await fetchJson(this.fetchImpl, url, { headers: { Authorization: `Bearer ${token}` } });
    return data?.value || [];
  }

  async listLondonInbox(limit = 10) {
    if (!this.londonMailbox) throw new Error('LONDON_MINACO_EMAIL is not configured.');
    return this.#listInbox(this.londonMailbox, limit);
  }

  async listPrincipalInbox(limit = 5) {
    if (!this.ramyMailbox) throw new Error('RAMY_MINACO_EMAIL is not configured.');
    return this.#listInbox(this.ramyMailbox, clampLimit(limit, 5, 10));
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

  async listPrincipalCalendar({ startIso, endIso, limit = 20 } = {}) {
    if (!this.ramyMailbox) throw new Error('RAMY_MINACO_EMAIL is not configured.');
    if (!startIso || !endIso) throw new Error('Calendar start and end are required.');

    const start = new Date(startIso);
    const end = new Date(endIso);
    if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime()) || end <= start) {
      throw new Error('Calendar start/end range is invalid.');
    }

    const token = await this.#getToken(this.readCreds, this.readToken);
    const url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.ramyMailbox)}/calendarView`);
    url.searchParams.set('startDateTime', start.toISOString());
    url.searchParams.set('endDateTime', end.toISOString());
    url.searchParams.set('$top', String(clampLimit(limit, 20, 50)));
    url.searchParams.set('$select', 'id,subject,start,end,location,organizer,isCancelled,isAllDay');
    url.searchParams.set('$orderby', 'start/dateTime');

    const data = await fetchJson(this.fetchImpl, url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.timezone="Eastern Standard Time"',
      },
    });
    return data?.value || [];
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
