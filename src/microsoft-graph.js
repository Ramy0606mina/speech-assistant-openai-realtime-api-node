import { fetchJson } from './http.js';
import { createHash } from 'node:crypto';

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
    return this.#listInbox(this.ramyMailbox, clampLimit(limit, 5, 25));
  }

  async listFollowUps() {
    const token=await this.#getToken(this.actionCreds,this.actionToken);
    const base=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.principalMailbox)}`;
    const headers={Authorization:`Bearer ${token}`,Prefer:'outlook.body-content-type="text"'};
    const calendars=await fetchJson(this.fetchImpl,`${base}/calendars?$select=id,name,owner&$top=100`,{headers});
    const matches=(calendars.value||[]).filter(c=>c.name==='London Action Register' && normalizeEmail(c.owner?.address)===this.principalMailbox);
    if(matches.length!==1)throw new Error('London Action Register unavailable.');
    const first=new URL(`${base}/calendars/${encodeURIComponent(matches[0].id)}/events?$top=100&$select=id,subject,body,isCancelled`);
    let next=first;const visited=new Set();const actions=[];
    while(next){
      if(next.origin!==first.origin || next.pathname!==first.pathname || visited.has(next.href) || visited.size>=100)throw new Error('Action register retrieval incomplete.');
      visited.add(next.href);
      const page=await fetchJson(this.fetchImpl,next,{headers});
      if(!Array.isArray(page.value))throw new Error('Action register response invalid.');
      for(const event of page.value){
        if(event.isCancelled)continue;
        const content=String(event.body?.content||'');
        if(!content.trim().startsWith('LONDON_ACTION_V1'))continue;
        const start=content.indexOf('{');const end=content.lastIndexOf('}');
        let action;try{action=JSON.parse(content.slice(start,end+1));}catch{throw new Error('An action record could not be read.');}
        if(['CLOSED','COMPLETED','CANCELLED','DONE'].includes(String(action.status).toUpperCase()))continue;
        actions.push({title:action.title,status:action.status,nextFollowUp:action.nextFollowUp,priority:action.priority,nextAction:action.nextAction});
      }
      next=page['@odata.nextLink']?new URL(page['@odata.nextLink']):null;
    }
    return actions;
  }

  async getLondonMessage(messageId) {
    if (!messageId) throw new Error('messageId is required.');
    const token = await this.#getToken(this.readCreds, this.readToken);
    const url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.londonMailbox)}/messages/${encodeURIComponent(messageId)}`);
    url.searchParams.set('$select', 'id,internetMessageId,conversationId,subject,from,sender,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,hasAttachments');
    return fetchJson(this.fetchImpl, url, {
      headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.body-content-type="text"' },
    });
  }

  async getLondonAttachments(messageId) {
    const token = await this.#getToken(this.readCreds, this.readToken);
    let url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.londonMailbox)}/messages/${encodeURIComponent(messageId)}/attachments`);
    const parts = [];
    let total = 0;
    let count = 0;
    while (url) {
      if (url.origin !== 'https://graph.microsoft.com') throw new Error('Invalid attachment pagination URL.');
      const data = await fetchJson(this.fetchImpl, url, { headers: { Authorization: `Bearer ${token}` } });
      for (const item of data.value || []) {
        if (item.isInline) continue;
        if (++count > 20) throw new Error('Too many attachments; maximum is 20.');
        const name = String(item.name || 'attachment');
        const supported = /\.(pdf|docx?|xlsx?|pptx?|txt|csv|md|rtf)$/i.test(name);
        if (item['@odata.type'] !== '#microsoft.graph.fileAttachment' || !supported) {
          parts.push({ type: 'input_text', text: `Attachment not analyzed (unsupported type): ${name}` });
          continue;
        }
        if (!item.contentBytes) throw new Error(`Attachment content unavailable: ${name}`);
        total += Buffer.from(item.contentBytes, 'base64').length;
        if (total > 20 * 1024 * 1024) throw new Error('Attachments exceed the 20 MB analysis limit.');
        parts.push({ type: 'input_file', filename: name, file_data: `data:${item.contentType || 'application/octet-stream'};base64,${item.contentBytes}` });
      }
      url = data['@odata.nextLink'] ? new URL(data['@odata.nextLink']) : null;
    }
    return parts;
  }

  async listPrincipalCalendar({ startIso, endIso, limit = 20 } = {}) {
    if (!this.ramyMailbox) throw new Error('RAMY_MINACO_EMAIL is not configured.');
    if (!startIso || !endIso) throw new Error('Calendar start and end are required.');

    const start = new Date(startIso);
    const end = new Date(endIso);
    if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime()) || end <= start) {
      throw new Error('Calendar start/end range is invalid.');
    }

    const calendarToken = await this.#getToken(this.actionCreds, this.actionToken);
    const url = new URL(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.ramyMailbox)}/calendarView`);
    url.searchParams.set('startDateTime', start.toISOString());
    url.searchParams.set('endDateTime', end.toISOString());
    url.searchParams.set('$top', String(clampLimit(limit, 20, 50)));
    url.searchParams.set('$select', 'id,subject,start,end,location,organizer,isCancelled,isAllDay,showAs');
    url.searchParams.set('$orderby', 'start/dateTime');

    const events = [];
    const visited = new Set();
    let next = url;
    while (next) {
      if (next.origin !== url.origin || next.pathname !== url.pathname || next.username || next.password) throw new Error('Calendar pagination address is invalid; availability was not established.');
      if (visited.has(next.href) || visited.size >= 200) throw new Error('Calendar retrieval did not complete; request a narrower date range.');
      visited.add(next.href);
      const data = await fetchJson(this.fetchImpl, next, { headers: {
        Authorization: `Bearer ${calendarToken}`, Prefer: 'outlook.timezone="Eastern Standard Time"',
      } });
      if (!Array.isArray(data?.value)) throw new Error('Calendar response is invalid; availability was not established.');
      events.push(...data.value);
      if (events.length > 10000) throw new Error('Calendar range is too large; request a narrower date range.');
      next = data['@odata.nextLink'] ? new URL(data['@odata.nextLink']) : null;
    }
    return events;
  }

  async createFollowUp({ title, date, notes = '', taskKey, reminder = true }) {
    if (typeof reminder !== 'boolean') throw new Error('Reminder preference must be true or false.');
    if (!this.principalMailbox || !taskKey || !String(title || '').trim() || String(title).length > 180) throw new Error('Follow-up title and source are required.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isFinite(Date.parse(date)) || new Date(date).toISOString().slice(0,10) !== date) throw new Error('Follow-up requires an explicit valid date.');
    if (String(notes).length > 4000) throw new Error('Follow-up notes are too long.');
    const token = await this.#getToken(this.actionCreds, this.actionToken);
    const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.principalMailbox)}`;
    const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
    const lists = await fetchJson(this.fetchImpl, `${base}/calendars?$select=id,name,owner&$top=100`, { headers });
    const matches = (lists.value || []).filter(c => c.name === 'London Action Register' && normalizeEmail(c.owner?.address) === this.principalMailbox);
    if (matches.length !== 1) throw new Error('A unique existing London Action Register was not found; no task was created.');
    const actionId = createHash('sha256').update(`${taskKey}\n${title}\n${date}`).digest('hex');
    const now = new Date().toISOString();
    const next = new Date(Date.parse(date)+86400000).toISOString().slice(0,10);
    const action = { actionId, title:String(title).trim(), owner:'London', dateOpened:now.slice(0,10), nextFollowUp:date,
      status:'ACTIVE',priority:'NORMAL',nextAction:String(title).trim(),source:'London owner email',notes:String(notes),reminder,createdAt:now,updatedAt:now };
    const result = await fetchJson(this.fetchImpl, `${base}/calendars/${encodeURIComponent(matches[0].id)}/events`, {
      method:'POST',headers,body:JSON.stringify({ subject:`[NORMAL] [ACTIVE] ${action.title}`, body:{contentType:'text',content:`LONDON_ACTION_V1\n${JSON.stringify(action,null,2)}`},
        start:{dateTime:`${date}T${reminder ? '09:00:00' : '00:00:00'}`,timeZone:'Eastern Standard Time'},end:{dateTime:reminder ? `${date}T09:15:00` : `${next}T00:00:00`,timeZone:'Eastern Standard Time'},
        isAllDay:!reminder,showAs:'free',sensitivity:'private',isReminderOn:reminder,reminderMinutesBeforeStart:0,attendees:[],transactionId:actionId }),
    });
    if (!result?.id) throw new Error('Microsoft did not confirm task creation; delivery requires review.');
    return { title:action.title,date,id:result.id,calendar:'London Action Register',reminder:reminder ? 'Outlook alert at 9 a.m. Eastern on the due date' : 'None' };
  }

  async sendMail({ to, subject, body, cc = [] }) {
    const addresses = (Array.isArray(to) ? to : [to]);
    const copies = (Array.isArray(cc) ? cc : [cc]).filter(Boolean);
    if (!this.principalMailbox || !addresses.length || addresses.some(address => normalizeEmail(address) !== this.principalMailbox) || copies.some(address => normalizeEmail(address) !== this.principalMailbox)) {
      throw new Error('Automatic email is restricted to the configured principal; other recipients are draft-only.');
    }
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
