import {createHash} from 'node:crypto';
import {formattedDraftBody} from './voice-gateway.js';

const hash = value => createHash('sha256').update(value).digest('hex');
const normalize = text => String(text || '').trim().replace(/\s+/g,' ');
const sendCommand = /^(?:(?:yes|please|go ahead)[, ]+)?send (?:it|the email|this email|the draft|this draft|the reply)(?: now| please)?[.!]?$/i;
export const draftOnlyReply = 'London only saves email drafts. Nothing was sent.';
const keepCommand = /^(?:(?:please|just) )?(?:keep|leave|save) (?:it|this|the email|the draft)(?: (?:as a draft|in (?:my |the |outlook )?drafts?(?: folder| box)?))?[.!]?$/i;
const noSendCommand = /^(?:(?:please|just) )?(?:do not|don't|don’t|never) send(?: it| the email| the draft)?[.!]?$/i;
const reviewCommand = /^(?:show|read|review|refresh)(?: me)? (?:it|the draft|this draft|the email)[.!]?$/i;
const reviseCommand = /^(?:(?:please|also|and) )?(?:add|include|mention|change|replace|remove|revise|rewrite|rephrase|translate|make|shorten|lengthen)\b/i;
const readRequest = /^(?:what|where|which|show|read|list|check|find|search|did|is|are)\b/i;

export function requestsEmailDraft(text) {
  const body = normalize(text);
  if (readRequest.test(body)) return false;
  if (/\b(do not|don't|don’t|never)\b[^.!?]*\b(draft|save|reply|respond|compose|write)\b/i.test(body)) return false;
  return /\b(draft|compose)\b|\b(?:write|prepare|save)\b.*\b(email|mail|drafts|reply)\b|^(?:please )?(?:email|reply|respond)\b|\b(?:reply|respond) to\b|\bsend (?:an? |new )?email\b/i.test(body);
}

export function draftVersion(message) {
  return hash(JSON.stringify({id:message.id,changeKey:message.changeKey,subject:message.subject,body:message.body,
    from:message.from,sender:message.sender,replyTo:message.replyTo,to:message.toRecipients,cc:message.ccRecipients,bcc:message.bccRecipients,hasAttachments:message.hasAttachments}));
}

export const updateSmsDraftTool = {
  type:'function',name:'update_saved_email_draft',strict:false,
  description:'Revise the body of the current saved Outlook draft in place. Preserve the recipient, subject and reply thread. Never sends. Supply the full revised professionally formatted body.',
  parameters:{type:'object',properties:{body:{type:'string'}},required:['body'],additionalProperties:false},
};

// SMS email actions only create or revise drafts. There is no sending action,
// including for conversations saved by versions that offered send confirmation.
export class SmsEmail {
  constructor({graph,dropbox,state,now=()=>new Date()}) { Object.assign(this,{graph,dropbox,state,now}); }
  get pending() { return this.state.state.smsConversation?.pendingEmail; }
  save(pendingEmail) {
    this.state.state.smsConversation={...this.state.state.smsConversation,pendingEmail};
    this.state.save();
  }
  cancelPending() { this.save(null); }
  validate(request) {
    if (!request.requestKey || !request.owner || request.owner.toLowerCase() !== this.graph.principalMailbox?.toLowerCase()) throw new Error('Authenticated SMS owner required.');
  }
  isCurrent(request) {
    const p=this.pending;
    const lastUser=request.history?.findLast(item=>item.role==='user');
    const lastReply=request.history?.at(-1);
    return Boolean(p && p.owner===request.owner && p.requestKey===lastUser?.requestKey && lastReply?.role==='assistant'
      && lastReply.content===p.reply && this.now().getTime()-Date.parse(p.updatedAt)<86400000);
  }
  async prepare(request) {
    this.validate(request);
    const text=normalize(request.body), current=this.isCurrent(request);
    if (sendCommand.test(text)) return {reply:draftOnlyReply};
    if (keepCommand.test(text) || noSendCommand.test(text)) {
      if (keepCommand.test(text) && current && this.pending?.status==='composing') return {create:true};
      if (!current || !this.pending?.id) return {reply:'Nothing was sent. Which email would you like me to save in Outlook Drafts?'};
      const message=await this.graph.getVoiceMessage(this.pending.mailbox,this.pending.id);
      if (message.isDraft !== true) return {reply:'That message is no longer a draft. I did not send anything.'};
      return {reply:this.remember(message,request,this.pending.mailbox)};
    }
    if (reviewCommand.test(text) && this.pending?.id) {
      const message=await this.graph.getVoiceMessage(this.pending.mailbox,this.pending.id);
      if (message.isDraft !== true) return {reply:'That message is no longer a draft. I did not send anything.'};
      return {reply:this.remember(message,request,this.pending.mailbox)};
    }
    if (current && this.pending?.id && reviseCommand.test(text)) {
      const message=await this.graph.getVoiceMessage(this.pending.mailbox,this.pending.id);
      if (message.isDraft !== true) return {reply:'That message is no longer a draft. I did not change or send it.'};
      return {revise:true,draft:message};
    }
    if (requestsEmailDraft(text)) {
      // A new request invalidates the earlier revision target, even if clarification
      // or a provider failure prevents the new draft from being saved.
      this.save({owner:request.owner,status:'composing',requestKey:request.requestKey,updatedAt:this.now().toISOString()});
      return {create:true};
    }
    if (current && this.pending?.status==='composing' && !readRequest.test(text)
      && !/^(?:thanks|thank you|ok|okay|perfect|great|no|never mind|cancel)[.!]?$/i.test(text)) return {create:true};
    return {};
  }
  recordAnswer(request,reply) {
    if (this.pending?.status==='composing') this.save({...this.pending,requestKey:request.requestKey,reply,updatedAt:this.now().toISOString()});
  }
  remember(message,request,mailbox='principal') {
    if (!message?.id || message.isDraft!==true) throw new Error('Microsoft did not confirm the saved draft.');
    const recipients=(message.toRecipients||[]).map(item=>item.emailAddress?.address).filter(Boolean).join(', ');
    const subject=String(message.subject||'(No subject)').replace(/\s+/g,' ').slice(0,90);
    const reply=`Saved in Outlook Drafts for ${this.graph.voiceMailbox(mailbox)}.\nTo: ${recipients.slice(0,150)||'(not set)'}\nSubject: ${subject}\nNothing was sent.`;
    this.save({owner:request.owner,id:message.id,mailbox,version:draftVersion(message),status:'saved',requestKey:request.requestKey,
      updatedAt:this.now().toISOString(),reply});
    return reply;
  }
  async saved(result,request,mailbox='principal') {
    if (!result?.id || result.isDraft!==true) throw new Error('No verified draft ID.');
    const message=await this.graph.getVoiceMessage(mailbox,result.id);
    if (message.id!==result.id) throw new Error('The saved draft could not be read back.');
    return this.remember(message,request,mailbox);
  }
  async update(args,request,expected) {
    const p=this.pending;
    if (!p?.id || !expected || !this.isCurrent(request)) throw new Error('No current draft to revise.');
    const body=formattedDraftBody(args.body);
    const live=await this.graph.getVoiceMessage(p.mailbox,p.id);
    if (live.isDraft!==true || draftVersion(live)!==draftVersion(expected)) throw new Error('Draft changed during revision.');
    const key='sms-draft-update-'+hash(request.owner+request.requestKey);
    if (!await this.dropbox.createDeliveryRecord(key,{status:'attempted',at:this.now().toISOString()})) throw new Error('Draft update already attempted.');
    this.save({...p,status:'updating'});
    const result=await this.graph.updateSmsDraft({mailbox:p.mailbox,id:p.id,body});
    return this.saved(result,request,p.mailbox);
  }
}
