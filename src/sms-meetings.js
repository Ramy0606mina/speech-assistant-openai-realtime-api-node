import {createHash} from 'node:crypto';
const hash=value=>createHash('sha256').update(value).digest('hex');
const codeFrom=reply=>/^Teams proposal:/.test(reply||'') ? reply.match(/CONFIRM MEETING ([a-f0-9]{12})/i)?.[1]?.toLowerCase() : null;
const plainReply=reply=>String(reply).replace(/Reply exactly:\s*CONFIRM MEETING [a-f0-9]{12}/i,'Reply confirm to send the invitation.');
const noPending='There is no current meeting proposal to confirm. Please send the meeting details again. No invitation was sent.';

// User-facing SMS contains no internal identifiers. The reference is bound to
// the authenticated owner and the actual proposal SMS, with a durable receipt.
export class SmsMeetings {
  constructor({meetings,state,graph,dropbox,sms,now=()=>new Date()}) {
    Object.assign(this,{meetings,state,graph,dropbox,sms,now});
  }
  get pending(){return this.state.state.smsConversation?.pendingMeeting;}
  save(pendingMeeting){
    this.state.state.smsConversation={...this.state.state.smsConversation,pendingMeeting};
    this.state.save();
  }
  async cancelPending(){
    if(this.pending?.code && this.pending.status==='pending')await this.dropbox.createDeliveryRecord('text-meeting-cancelled-'+this.pending.code,{owner:this.pending.owner});
    this.save({status:'changed'});
  }
  async recover(owner) {
    // Only the most recent actual outgoing SMS may restore a lost reference.
    // A newer clarification/cancellation/result makes older proposals unusable.
    const outgoing=await this.sms.latestOutgoing();
    if(!outgoing || !['sent','delivered'].includes(outgoing.status))return null;
    const reference=await this.dropbox.readDeliveryRecord('sms-meeting-reference-'+hash(outgoing.id));
    let code=reference?.code;
    if(reference && (reference.owner!==owner || reference.bodyHash!==hash(outgoing.body)))return null;
    if(!code)code=codeFrom(outgoing.body); // Migrate the last pre-upgrade proposal.
    if(!code)return null;
    const proposal=await this.dropbox.readDeliveryRecord('text-meeting-proposal-'+code);
    if(!proposal || proposal.owner!==owner || ![proposal.reply,plainReply(proposal.reply)].includes(outgoing.body))return null;
    const pending={code,owner,sourceRequestKey:proposal.requestKey,presentedAt:outgoing.sentAt,status:'pending'};
    this.save(pending);
    return pending;
  }
  async handle(request) {
    const text=String(request.text||'').trim();
    const owner=String(request.owner||'').toLowerCase();
    if(!owner || owner!==this.graph.principalMailbox?.toLowerCase() || !request.requestKey)throw Error('Authenticated owner and SMS source required.');
    if(/^confirm[.!]?$/i.test(text)) {
      const pending=this.pending || await this.recover(owner);
      if(!pending || pending.owner!==owner)return noPending;
      if(pending.status==='completed')return pending.reply;
      if(pending.status!=='pending')return 'This invitation was already attempted or the proposal changed. Check your calendar or send fresh meeting details; no duplicate invitation will be sent.';
      if(!Number.isFinite(Date.parse(pending.presentedAt)) || !Number.isFinite(Date.parse(request.receivedAt)) || Date.parse(request.receivedAt)<Math.floor(Date.parse(pending.presentedAt)/1000)*1000 || pending.sourceRequestKey===request.requestKey)return noPending;
      this.save({...pending,status:'attempted'});
      let reply;
      try { reply=await this.meetings.confirm(pending.code,owner,request.requestKey,request.maxReplyLength||480); }
      catch { reply='The invitation could not be verified. Check your calendar before requesting it again; I will not automatically retry it.'; }
      this.save({...pending,status:/^(?:Teams invitation submitted:|This Teams invitation was already submitted)/.test(reply)?'completed':'blocked',reply});
      return reply;
    }
    // Explicitly supplied old codes are no longer the SMS confirmation UI.
    if(/^confirm meeting\b/i.test(text))return this.pending?.status==='pending' ? 'To send the current meeting invitation, reply confirm.' : noPending;
    const reply=await this.meetings.handle({...request,allowImplicitConflictConfirmation:false});
    if(!reply)return null;
    const code=codeFrom(reply);
    if(this.pending?.code && this.pending.code!==code && this.pending.status==='pending') {
      await this.dropbox.createDeliveryRecord('text-meeting-cancelled-'+this.pending.code,{owner});
    }
    if(!code){this.save({owner,status:'changed'});return reply;}
    const proposal=await this.dropbox.readDeliveryRecord('text-meeting-proposal-'+code);
    if(!proposal || proposal.owner!==owner || proposal.reply!==reply)throw Error('Unverified SMS meeting proposal.');
    this.save({owner,code,sourceRequestKey:request.requestKey,status:'pending',presentedAt:null});
    return plainReply(reply);
  }
  async recordReply({requestKey,replyId,text,sentAt}) {
    const pending=this.pending;
    if(!pending || pending.status!=='pending' || pending.sourceRequestKey!==requestKey)return;
    this.save({...pending,presentedAt:sentAt});
    await this.dropbox.createDeliveryRecord('sms-meeting-reference-'+hash(replyId),{...pending,presentedAt:sentAt,bodyHash:hash(text)});
  }
}
