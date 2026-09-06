import {fetchJson} from './http.js';
import {normalizePhone} from './voice-gateway.js';
export class SmsClient {
  constructor({accountSid,authToken,from,to,fetchImpl=fetch}) {Object.assign(this,{accountSid,authToken,from:normalizePhone(from),to:normalizePhone(to),fetchImpl});}
  get configured(){return Boolean(/^AC[\da-f]{32}$/i.test(this.accountSid||'') && this.authToken && /^\+\d{10,15}$/.test(this.from) && /^\+\d{10,15}$/.test(this.to));}
  async send(body){
    if(!this.configured)throw new Error('London SMS credentials are not configured.');
    if(typeof body!=='string' || !body.trim() || body.length>480)throw new Error('SMS must contain 1–480 characters.');
    const result=await fetchJson(this.fetchImpl,`https://api.twilio.com/2010-04-01/Accounts/${this.accountSid}/Messages.json`,{method:'POST',headers:{Authorization:`Basic ${Buffer.from(`${this.accountSid}:${this.authToken}`).toString('base64')}`,'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({From:this.from,To:this.to,Body:body.trim()}).toString()});
    if(!result?.sid || ['failed','undelivered','canceled'].includes(result.status))throw new Error('Twilio did not accept the SMS.');
    return {accepted:true,status:result.status||'accepted',id:result.sid};
  }
}
