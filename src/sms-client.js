import {fetchJson} from './http.js';
import {normalizePhone} from './voice-gateway.js';
export class SmsClient {
  constructor({accountSid,authToken,from,to,fetchImpl=fetch}) {Object.assign(this,{accountSid,authToken,from:normalizePhone(from),to:normalizePhone(to),fetchImpl});}
  get configured(){return Boolean(/^AC[\da-f]{32}$/i.test(this.accountSid||'') && this.authToken && /^\+\d{10,15}$/.test(this.from) && /^\+\d{10,15}$/.test(this.to));}
  async configureInbound(publicUrl) {
    const origin = new URL(publicUrl);
    if (!this.configured || origin.protocol !== 'https:' || origin.username || origin.password) throw new Error('Invalid SMS webhook configuration.');
    const desired = `${origin.origin}/incoming-sms`;
    const base = `https://api.twilio.com/2010-04-01/Accounts/${this.accountSid}/IncomingPhoneNumbers`;
    const headers = { Authorization: `Basic ${Buffer.from(`${this.accountSid}:${this.authToken}`).toString('base64')}` };
    const listed = await fetchJson(this.fetchImpl, `${base}.json?${new URLSearchParams({PhoneNumber:this.from})}`, {headers});
    const matches = (listed.incoming_phone_numbers || []).filter(item => item.phone_number === this.from);
    if (matches.length !== 1 || !/^PN[\da-f]{32}$/i.test(matches[0].sid) || matches[0].capabilities?.sms !== true) throw new Error('Exactly one SMS-capable London number is required.');
    const number = matches[0];
    if (number.sms_application_sid) throw new Error('London SMS uses a TwiML application; its routing needs review.');
    if (number.sms_url !== desired || number.sms_method !== 'POST' || number.sms_fallback_url !== desired) {
      await fetchJson(this.fetchImpl, `${base}/${number.sid}.json`, {method:'POST',headers:{...headers,'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({SmsUrl:desired,SmsMethod:'POST',SmsFallbackUrl:desired,SmsFallbackMethod:'POST'}).toString()});
    }
    const verified = await fetchJson(this.fetchImpl, `${base}/${number.sid}.json`, {headers});
    if (verified.sms_url !== desired || verified.sms_method !== 'POST' || verified.sms_fallback_url !== desired || verified.sms_application_sid) throw new Error('SMS webhook was not confirmed.');
    return 'number-webhook-verified';
  }
  async listIncoming(since) {
    if (!this.configured || !Number.isFinite(since)) throw new Error('Incoming SMS is not configured.');
    const path = `/2010-04-01/Accounts/${this.accountSid}/Messages.json`;
    const query = new URLSearchParams({ From: this.to, To: this.from, 'DateSent>': new Date(since - 86400000).toISOString().slice(0, 10), PageSize: '100' });
    let url = `https://api.twilio.com${path}?${query}`;
    const messages = [];
    for (let page = 0; page < 20; page++) {
      const result = await fetchJson(this.fetchImpl, url, { headers: { Authorization: `Basic ${Buffer.from(`${this.accountSid}:${this.authToken}`).toString('base64')}` } });
      if (!Array.isArray(result?.messages)) throw new Error('Invalid incoming SMS response.');
      for (const item of result.messages) {
        const received = Date.parse(item.date_created || item.date_sent);
        if (item.direction !== 'inbound' || item.from !== this.to || item.to !== this.from || !/^(SM|MM)[\da-f]{32}$/i.test(item.sid || '') || !Number.isFinite(received) || received < since) continue;
        messages.push({ sid: item.sid, body: item.body, receivedAt: new Date(received).toISOString(), numMedia: Number(item.num_media || 0) });
      }
      if (!result.next_page_uri) return messages.sort((a, b) => a.receivedAt.localeCompare(b.receivedAt) || a.sid.localeCompare(b.sid));
      const next = new URL(result.next_page_uri, 'https://api.twilio.com');
      if (next.origin !== 'https://api.twilio.com' || next.pathname !== path) throw new Error('Invalid SMS pagination URL.');
      url = next.href;
    }
    throw new Error('Incoming SMS scan limit reached.');
  }
  async messageStatus(sid) {
    if (!this.configured || !/^SM[\da-f]{32}$/i.test(sid || '')) throw new Error('Invalid SMS id.');
    const result = await fetchJson(this.fetchImpl, `https://api.twilio.com/2010-04-01/Accounts/${this.accountSid}/Messages/${sid}.json`, { headers: { Authorization: `Basic ${Buffer.from(`${this.accountSid}:${this.authToken}`).toString('base64')}` } });
    return result.status || 'unknown';
  }
  async latestOutgoing() {
    if(!this.configured)throw Error('SMS is not configured.');
    const url=`https://api.twilio.com/2010-04-01/Accounts/${this.accountSid}/Messages.json?${new URLSearchParams({From:this.from,To:this.to,PageSize:'1'})}`;
    const result=await fetchJson(this.fetchImpl,url,{headers:{Authorization:`Basic ${Buffer.from(`${this.accountSid}:${this.authToken}`).toString('base64')}`}});
    const item=result.messages?.[0];
    if(!item || item.from!==this.from || item.to!==this.to || !String(item.direction).startsWith('outbound') || !/^SM[\da-f]{32}$/i.test(item.sid||''))return null;
    return {id:item.sid,body:String(item.body||''),status:item.status,sentAt:item.date_sent||item.date_created};
  }
  async send(body){
    if(!this.configured)throw new Error('London SMS credentials are not configured.');
    if(typeof body!=='string' || !body.trim() || body.length>480)throw new Error('SMS must contain 1–480 characters.');
    const result=await fetchJson(this.fetchImpl,`https://api.twilio.com/2010-04-01/Accounts/${this.accountSid}/Messages.json`,{method:'POST',headers:{Authorization:`Basic ${Buffer.from(`${this.accountSid}:${this.authToken}`).toString('base64')}`,'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({From:this.from,To:this.to,Body:body.trim()}).toString()});
    if(!result?.sid || ['failed','undelivered','canceled'].includes(result.status))throw new Error('Twilio did not accept the SMS.');
    return {accepted:true,status:result.status||'accepted',id:result.sid};
  }
}
