import { fetchJson } from './http.js';
import { createHash } from 'node:crypto';

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function clampLimit(value, fallback = 10, max = 25) {
  return Math.min(max, Math.max(1, Number(value) || fallback));
}

function nameWords(value) {
  return String(value||'').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim().split(/\s+/).filter(Boolean);
}

function editDistance(a,b) {
  const row=Array.from({length:b.length+1},(_,i)=>i);
  for(let i=1;i<=a.length;i++){
    let diagonal=row[0];row[0]=i;
    for(let j=1;j<=b.length;j++){
      const previous=row[j];
      row[j]=Math.min(row[j]+1,row[j-1]+1,diagonal+(a[i-1]===b[j-1]?0:1));
      diagonal=previous;
    }
  }
  return row[b.length];
}

function approximatePersonScore(query,sender) {
  const q=nameWords(query),s=nameWords(sender);
  if(q.length<2||q.length>4||s.length<2)return 0;
  const pairs=[[q[0],s[0]],[q[q.length-1],s[s.length-1]]];
  const scores=pairs.map(([a,b])=>1-editDistance(a,b)/Math.max(a.length,b.length,1));
  return scores[1]>=0.45 ? (scores[0]+scores[1])/2 : 0;
}

const MEETING_ZONE_ALIASES=new Map([
  ['america/toronto',{iana:'America/Toronto',graph:'Eastern Standard Time',label:'America/Toronto'}],
  ['america/new_york',{iana:'America/Toronto',graph:'Eastern Standard Time',label:'America/Toronto'}],
  ['eastern standard time',{iana:'America/Toronto',graph:'Eastern Standard Time',label:'America/Toronto'}],
  ['eastern time',{iana:'America/Toronto',graph:'Eastern Standard Time',label:'America/Toronto'}],
  ['et',{iana:'America/Toronto',graph:'Eastern Standard Time',label:'America/Toronto'}],
  ['utc',{iana:'UTC',graph:'UTC',label:'UTC'}],
  ['america/chicago',{iana:'America/Chicago',graph:'Central Standard Time',label:'America/Chicago'}],
  ['central standard time',{iana:'America/Chicago',graph:'Central Standard Time',label:'America/Chicago'}],
  ['america/denver',{iana:'America/Denver',graph:'Mountain Standard Time',label:'America/Denver'}],
  ['mountain standard time',{iana:'America/Denver',graph:'Mountain Standard Time',label:'America/Denver'}],
  ['america/los_angeles',{iana:'America/Los_Angeles',graph:'Pacific Standard Time',label:'America/Los_Angeles'}],
  ['pacific standard time',{iana:'America/Los_Angeles',graph:'Pacific Standard Time',label:'America/Los_Angeles'}],
]);

function meetingZone(value) {
  const raw=String(value||'').trim();
  const alias=MEETING_ZONE_ALIASES.get(raw.toLowerCase());
  const zone=alias||{iana:raw,graph:raw,label:raw};
  try { new Intl.DateTimeFormat('en-CA',{timeZone:zone.iana}).format(new Date()); }
  catch { throw new Error('Meeting timezone is unsupported. Use America/Toronto unless another timezone is explicitly requested.'); }
  return zone;
}

function zonedDateTime(date,iana) {
  const parts=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:iana,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(date).filter(p=>p.type!=='literal').map(p=>[p.type,p.value]));
  return `${parts.year}-${parts.month}-${parts.day}T${parts.hour}:${parts.minute}:${parts.second}`;
}

function verifiedMeetingStart(startIso,timezone) {
  const raw=String(startIso||'');
  const match=raw.match(/^(\d{4}-\d{2}-\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(Z|[+-]\d{2}:\d{2})$/);
  const date=new Date(raw);const zone=meetingZone(timezone||'America/Toronto');
  if(!match||!Number.isFinite(date.getTime()))throw new Error('Meeting start requires an explicit ISO date, time, and offset.');
  const supplied=`${match[1]}T${match[2]}:${match[3]}:${match[4]||'00'}`;
  if(zonedDateTime(date,zone.iana)!==supplied)throw new Error('Meeting time offset does not match the requested timezone or its daylight-saving rules.');
  return {date,zone,local:supplied};
}

function contactNameScore(query,name,address='') {
  const emailQuery=normalizeEmail(query);
  if(emailQuery.includes('@'))return emailQuery===normalizeEmail(address)?1:0;
  const q=nameWords(query),n=nameWords(name);
  if(q.length<1||n.length<1)return 0;
  if(q.join(' ')===n.join(' '))return 1;
  if(q.every(word=>n.includes(word)))return 0.92;
  return approximatePersonScore(query,name);
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

  async searchVoiceMessages({mailbox='principal',folder='inbox',query,startIso,endIso,maxScan=2000,maxResults=25}={}) {
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
    const approximate=[];
    const visited=new Set();
    const ownerPath=owner.toLowerCase();
    const allowedPaths=folder==='all'
      ? new Set([`/v1.0/users/${ownerPath}/messages`])
      : new Set([
          `/v1.0/users/${ownerPath}/mailfolders/${folder}/messages`,
          `/v1.0/users/${ownerPath}/mailfolders('${folder}')/messages`,
        ]);
    while(next&&scanned<maxScan) {
      let pagingPath='';
      try { pagingPath=decodeURIComponent(next.pathname).toLowerCase(); } catch {}
      if(next.origin!=='https://graph.microsoft.com'||next.username||next.password||!allowedPaths.has(pagingPath)||visited.has(next.href)||visited.size>=50)throw new Error('Email search paging was invalid; results were not established.');
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
        else {
          const sender=message.from?.emailAddress?.name||'';
          const score=approximatePersonScore(query,sender);
          if(score>=0.56)approximate.push({message,sender,score});
        }
      }
      next=page['@odata.nextLink']?new URL(page['@odata.nextLink']):null;
    }
    if(!matches.length&&approximate.length){
      approximate.sort((a,b)=>b.score-a.score||new Date(b.message.receivedDateTime||0)-new Date(a.message.receivedDateTime||0));
      const suggestedSender=approximate[0].sender;
      const selected=approximate.filter(item=>item.sender===suggestedSender).slice(0,maxResults).map(item=>item.message);
      return {messages:selected,scanned,complete:!next,approximateMatch:true,suggestedSender};
    }
    return {messages:matches,scanned,complete:!next};
  }

  async resolveVoiceContact({mailbox='principal',query,context='',maxMessages=5000}={}) {
    const owner=this.voiceMailbox(mailbox);
    const term=String(query||'').trim();
    if(!term||term.length>160)throw new Error('A contact name or email address is required.');
    const token=await this.#getToken(this.actionCreds,this.actionToken);
    const base=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}`;
    const headers={Authorization:`Bearer ${token}`,Prefer:'IdType="ImmutableId", outlook.body-content-type="text"'};
    const inbox=await fetchJson(this.fetchImpl,`${base}/mailFolders/inbox?$select=id,displayName,parentFolderId,childFolderCount`,{headers});
    if(!inbox?.id)throw new Error('The principal Inbox could not be opened for contact lookup.');
    const folders=[{id:inbox.id,name:inbox.displayName||'Inbox',path:inbox.displayName||'Inbox',childFolderCount:Number(inbox.childFolderCount||0)}];
    const queue=[folders[0]];const folderIds=new Set([String(inbox.id)]);
    const validPage=(url,suffix)=>{
      let path='';try{path=decodeURIComponent(url.pathname).toLowerCase();}catch{}
      return url.origin==='https://graph.microsoft.com'&&!url.username&&!url.password&&path.startsWith(`/v1.0/users/${owner.toLowerCase()}/`)&&path.endsWith(suffix);
    };
    while(queue.length){
      const parent=queue.shift();if(!parent.childFolderCount)continue;
      let next=new URL(`${base}/mailFolders/${encodeURIComponent(parent.id)}/childFolders?$top=100&$select=id,displayName,parentFolderId,childFolderCount`);const seen=new Set();
      while(next){
        if(!validPage(next,'/childfolders')||seen.has(next.href)||seen.size>=20)throw new Error('Mailbox folder lookup did not complete safely.');
        seen.add(next.href);const page=await fetchJson(this.fetchImpl,next,{headers});
        if(!Array.isArray(page?.value))throw new Error('Mailbox folder response was invalid.');
        for(const item of page.value){
          if(!item?.id||folderIds.has(String(item.id)))continue;
          if(folders.length>=200)throw new Error('Mailbox has too many nested folders for reliable contact lookup.');
          folderIds.add(String(item.id));const folder={id:item.id,name:item.displayName||'',path:`${parent.path}/${item.displayName||''}`,childFolderCount:Number(item.childFolderCount||0)};folders.push(folder);queue.push(folder);
        }
        next=page['@odata.nextLink']?new URL(page['@odata.nextLink']):null;
      }
    }
    const contextNeedle=nameWords(context).join(' ');const queryNeedle=nameWords(term).join(' ');
    const candidates=new Map();let scanned=0;
    for(const folder of folders){
      const folderText=nameWords(folder.path).join(' ');const relevant=Boolean((queryNeedle&&folderText.includes(queryNeedle))||(contextNeedle&&folderText.includes(contextNeedle)));
      let pages=0;let next=new URL(`${base}/mailFolders/${encodeURIComponent(folder.id)}/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime`);const seen=new Set();
      while(next&&scanned<maxMessages&&(pages<1||relevant&&pages<5)){
        if(!validPage(next,'/messages')||seen.has(next.href))throw new Error('Contact evidence paging did not complete safely.');
        seen.add(next.href);pages++;const page=await fetchJson(this.fetchImpl,next,{headers});
        if(!Array.isArray(page?.value))throw new Error('Contact evidence response was invalid.');
        for(const message of page.value){
          if(scanned++>=maxMessages)break;
          const people=[message.from?.emailAddress,...(message.toRecipients||[]).map(v=>v.emailAddress),...(message.ccRecipients||[]).map(v=>v.emailAddress)].filter(Boolean);
          for(const person of people){
            const address=normalizeEmail(person.address);if(!address||address===owner||address===this.readMailbox)continue;
            const name=String(person.name||'').trim();const baseScore=contactNameScore(term,name,address);if(baseScore<0.56)continue;
            const subjectContext=contextNeedle&&nameWords(message.subject).join(' ').includes(contextNeedle)?0.04:0;const score=Math.min(1,baseScore+(relevant?0.04:0)+subjectContext);
            const existing=candidates.get(address)||{address,name,score:0,evidenceCount:0,latestDateTime:'',folders:new Set()};existing.score=Math.max(existing.score,score);existing.evidenceCount++;existing.name=existing.name||name;const when=message.receivedDateTime||message.sentDateTime||'';if(when>existing.latestDateTime)existing.latestDateTime=when;existing.folders.add(folder.path);candidates.set(address,existing);
          }
        }
        next=page['@odata.nextLink']?new URL(page['@odata.nextLink']):null;
      }
    }
    const contacts=[...candidates.values()].map(c=>({...c,score:Number(c.score.toFixed(3)),folders:[...c.folders].slice(0,5)})).sort((a,b)=>b.score-a.score||b.evidenceCount-a.evidenceCount||b.latestDateTime.localeCompare(a.latestDateTime));
    if(!contacts.length)return {status:'not_found',contacts:[],foldersSearched:folders.length,messagesScanned:scanned};
    const resolved=!contacts[1]||contacts[0].score-contacts[1].score>=0.08;
    return {status:resolved?'resolved':'ambiguous',contacts:contacts.slice(0,resolved?1:5),foldersSearched:folders.length,messagesScanned:scanned};
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

  previewVoiceMeeting({startIso,durationMinutes,timezone='America/Toronto'}={}) {
    const verified=verifiedMeetingStart(startIso,timezone);const duration=Number(durationMinutes);
    if(!Number.isInteger(duration)||duration<15||duration>480)throw new Error('Meeting duration must be between 15 minutes and 8 hours.');
    const end=new Date(verified.date.getTime()+duration*60000);
    return {startIso:verified.date.toISOString(),endIso:end.toISOString(),localStart:verified.local,localEnd:zonedDateTime(end,verified.zone.iana),timezone:verified.zone.label,microsoftTimeZone:verified.zone.graph};
  }

  async createVoiceMeeting({title,startIso,durationMinutes,timezone='America/Toronto',attendees,body='',location='',onlineMeeting=false,transactionId}={}) {
    const subject=String(title||'').trim();
    const preview=this.previewVoiceMeeting({startIso,durationMinutes,timezone});const start=new Date(preview.startIso);
    const duration=Number(durationMinutes);
    const addresses=Array.isArray(attendees)?[...new Set(attendees.map(normalizeEmail))]:[];
    if(!this.principalMailbox||!subject||subject.length>180)throw new Error('Meeting title is required and must be at most 180 characters.');
    if(!addresses.length||addresses.length>20||addresses.some(value=>!/^[^\s<>@,;]+@[^\s<>@,;]+\.[^\s<>@,;]+$/.test(value)))throw new Error('Meeting invitations require one to twenty explicit attendee email addresses.');
    if(String(body).length>4000||String(location).length>300)throw new Error('Meeting notes or location are too long.');
    if(!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(transactionId||'')))throw new Error('Meeting transaction is invalid.');
    const end=new Date(start.getTime()+duration*60000);
    const token=await this.#getToken(this.actionCreds,this.actionToken);
    const eventUrl=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.principalMailbox)}/events`;
    const result=await fetchJson(this.fetchImpl,eventUrl,{
      method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({
        subject,body:{contentType:'Text',content:String(body)},start:{dateTime:preview.localStart,timeZone:preview.microsoftTimeZone},end:{dateTime:preview.localEnd,timeZone:preview.microsoftTimeZone},
        location:{displayName:String(location)},attendees:addresses.map(address=>({emailAddress:{address},type:'required'})),allowNewTimeProposals:true,transactionId,
        ...(onlineMeeting?{isOnlineMeeting:true,onlineMeetingProvider:'teamsForBusiness'}:{}),
      }),
    });
    if(!result?.id)throw new Error('Microsoft did not confirm meeting creation; invitation delivery was not established.');
    let verified=result;
    if(onlineMeeting&&(!result.isOnlineMeeting||!result.onlineMeeting?.joinUrl)){
      verified=await fetchJson(this.fetchImpl,`${eventUrl}/${encodeURIComponent(result.id)}?$select=id,isOnlineMeeting,onlineMeetingProvider,onlineMeeting`,{
        headers:{Authorization:`Bearer ${token}`},
      });
    }
    const joinLinkCreated=Boolean(verified?.isOnlineMeeting&&verified?.onlineMeetingProvider==='teamsForBusiness'&&verified?.onlineMeeting?.joinUrl);
    return {id:result.id,title:subject,...preview,durationMinutes:duration,attendees:addresses,location:String(location),calendar:'Primary Outlook calendar',invitationsSubmitted:true,onlineMeeting:Boolean(onlineMeeting),onlineMeetingProvider:onlineMeeting?'teamsForBusiness':'unknown',joinLinkCreated};
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

