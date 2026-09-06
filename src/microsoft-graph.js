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

  voiceMailbox(choice = 'principal') {
    const mailbox=choice==='principal' ? this.principalMailbox : choice==='london' ? this.readMailbox : '';
    if(!mailbox)throw new Error('Choose a connected mailbox: principal or london.');
    return mailbox;
  }

  async voiceRequest(mailbox,path,options={}) {
    const owner=this.voiceMailbox(mailbox);
    const token=await this.#getToken(this.actionCreds,this.actionToken);
    return fetchJson(this.fetchImpl,`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}${path}`,{
      ...options,headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json',Prefer:'IdType="ImmutableId", outlook.body-content-type="text"'},
    });
  }

  async listVoiceMessages(mailbox='principal',folder='inbox',limit=10) {
    if(!['inbox','drafts'].includes(folder))throw new Error('Unsupported mail folder.');
    const data=await this.voiceRequest(mailbox,`/mailFolders/${folder}/messages?$top=${clampLimit(limit)}&$orderby=lastModifiedDateTime desc&$select=id,subject,from,toRecipients,bodyPreview,isDraft,receivedDateTime,hasAttachments`);
    return data?.value||[];
  }

  async searchVoiceMessages({mailbox='principal',folder='inbox',query,startIso,endIso,maxScan=1000,maxResults=25}={}) {
    const owner=this.voiceMailbox(mailbox);
    const folders=['all','inbox','drafts','sentitems','deleteditems'];
    if(!folders.includes(folder))throw new Error('Unsupported mail folder.');
    const needle=String(query||'').trim().toLowerCase();
    if(!needle||needle.length>200)throw new Error('A sender, email address, subject, or phrase is required.');
    const start=startIso ? new Date(startIso) : null;
    const end=endIso ? new Date(endIso) : null;
    if((start && !Number.isFinite(start.getTime()))||(end && !Number.isFinite(end.getTime()))||(start&&end&&end<=start))throw new Error('The email date range is invalid.');
    maxScan=Math.min(2000,Math.max(50,Number(maxScan)||1000));
    maxResults=Math.min(50,Math.max(1,Number(maxResults)||25));
    const token=await this.#getToken(this.actionCreds,this.actionToken);
    const base=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}`;
    const collection=folder==='all' ? `${base}/messages` : `${base}/mailFolders/${folder}/messages`;
    const first=new URL(collection);
    first.searchParams.set('$top','50');
    first.searchParams.set('$select','id,subject,from,toRecipients,bodyPreview,isDraft,receivedDateTime,sentDateTime,hasAttachments');
    first.searchParams.set('$orderby','receivedDateTime desc');
    let next=first,scanned=0;
    const matches=[];
    const visited=new Set();
    while(next&&scanned<maxScan) {
      if(next.origin!=='https://graph.microsoft.com'||!next.pathname.startsWith(`/v1.0/users/${encodeURIComponent(owner)}/`)||!next.pathname.endsWith('/messages')||visited.has(next.href)||visited.size>=50)throw new Error('Email search paging was invalid; results were not established.');
      visited.add(next.href);
      const page=await fetchJson(this.fetchImpl,next,{headers:{Authorization:`Bearer ${token}`,Prefer:'IdType="ImmutableId", outlook.body-content-type="text"'}});
      if(!Array.isArray(page?.value))throw new Error('Email search response was invalid.');
      for(const message of page.value) {
        if(scanned++>=maxScan)break;
        const when=new Date(message.receivedDateTime||message.sentDateTime||0);
        if(start&&when<start)continue;
        if(end&&when>=end)continue;
        const haystack=[message.from?.emailAddress?.name,message.from?.emailAddress?.address,message.subject,message.bodyPreview].map(v=>String(v||'').toLowerCase()).join('\n');
        if(haystack.includes(needle)&&matches.length<maxResults)matches.push(message);
      }
      next=page['@odata.nextLink']?new URL(page['@odata.nextLink']):null;
    }
    return {messages:matches,scanned,complete:!next};
  }

  async getVoiceMessage(mailbox,id) {
    if(typeof id!=='string'||!id||id.length>2000)throw new Error('A message selected from the mailbox is required.');
    return this.voiceRequest(mailbox,`/messages/${encodeURIComponent(id)}?$select=id,subject,from,replyTo,toRecipients,ccRecipients,body,isDraft,hasAttachments`);
  }

  async createVoiceDraft({mailbox='principal',to=[],subject,body,messageId}) {
    if(typeof body!=='string'||!body.trim()||body.length>20000)throw new Error('A draft body of at most 20,000 characters is required.');
    let result;
    if(messageId) {
      const source=await this.getVoiceMessage(mailbox,messageId);
      if(source.isDraft)throw new Error('Choose a received message to reply to.');
      result=await this.voiceRequest(mailbox,`/messages/${encodeURIComponent(messageId)}/createReply`,{method:'POST',body:JSON.stringify({comment:body})});
    } else {
      if(typeof subject!=='string'||!subject.trim()||subject.length>250)throw new Error('A draft subject is required.');
      if(!Array.isArray(to)||!to.length||to.length>10||to.some(v=>typeof v!=='string'||v.length>254||!/^[^\s<>@,;]+@[^\s<>@,;]+\.[^\s<>@,;]+$/.test(v)))throw new Error('Provide explicit recipient email addresses; names must not be guessed.');
      result=await this.voiceRequest(mailbox,'/messages',{method:'POST',body:JSON.stringify({subject,body:{contentType:'Text',content:body},toRecipients:to.map(address=>({emailAddress:{address}}))})});
    }
    if(!result?.id||result.isDraft!==true)throw new Error('Microsoft did not confirm the saved draft. Check Drafts before retrying.');
    return {id:result.id,subject:result.subject,isDraft:true,mailbox:this.voiceMailbox(mailbox),folder:'Drafts',sent:false};
  }

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
