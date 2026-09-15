import {createHash} from 'node:crypto';

const hash = value => createHash('sha256').update(value).digest('hex');
const plain = reply => String(reply).replace(/Reply exactly:\s*CONFIRM MEETING [a-f0-9]{12}/i, 'Reply confirm to send the invitation.');
const noPending = 'There is no current meeting proposal in this email conversation. Please send the meeting details in this thread.';

// The visible reply is plain "confirm". The internal proposal stays bound to
// the authenticated owner, conversation and successfully sent proposal.
export class EmailMeetings {
  constructor({meetings, state, graph, dropbox, now = () => new Date()}) {
    Object.assign(this, {meetings, state, graph, dropbox, now});
  }
  key(owner, conversationId) { return hash(JSON.stringify([owner, conversationId])); }
  pending(key) { return this.state.state.emailMeetings?.[key]; }
  save(key, value) {
    this.state.state.emailMeetings ||= {};
    this.state.state.emailMeetings[key] = value;
    this.state.save();
  }
  async handle(request) {
    const {conversationId, requestKey, receivedAt} = request;
    const owner = String(request.owner || '').toLowerCase();
    const text = String(request.text || '').trim();
    if (!owner || owner !== this.graph.principalMailbox || !requestKey) throw Error('Authenticated owner and email source required.');
    if (!conversationId) return /^confirm[.!]?$/i.test(text) ? noPending : null;
    const key = this.key(owner, conversationId), pending = this.pending(key);
    if (/^confirm[.!]?$/i.test(text)) {
      if (!pending || pending.owner !== owner || pending.sourceRequestKey === requestKey) return noPending;
      if (pending.status === 'completed') return pending.reply;
      if (pending.status !== 'pending') return 'That proposal was changed or its invitation was already attempted. Please check your calendar; no duplicate invitation will be sent.';
      if (!Number.isFinite(Date.parse(receivedAt)) || !Number.isFinite(Date.parse(pending.presentedAt)) || Date.parse(receivedAt) < Math.floor(Date.parse(pending.presentedAt) / 1000) * 1000) return noPending;
      this.save(key, {...pending, status: 'attempted'});
      let reply;
      try { reply = await this.meetings.confirm(pending.code, owner, requestKey, Infinity); }
      catch { reply = 'The invitation could not be verified. Check your calendar before requesting it again; I will not automatically retry it.'; }
      this.save(key, {...pending, status: /^(?:Teams|In-person) invitation submitted:/.test(reply) ? 'completed' : 'blocked', reply});
      return reply;
    }
    if (/^confirm meeting\b/i.test(text)) return pending?.status === 'pending' ? 'To send the current meeting invitation, reply confirm.' : noPending;
    // A new request or correction supersedes the old proposal in this thread.
    const history = pending?.history || [];
    const reply = await this.meetings.handle({...request, history, allowImplicitConflictConfirmation: false});
    if (!reply) return null;
    const code = /^(?:Teams|In-person) proposal:/.test(reply) ? reply.match(/CONFIRM MEETING ([a-f0-9]{12})/i)?.[1] : null;
    if (pending?.code && pending.code !== code && ['pending', 'prepared'].includes(pending.status)) {
      await this.dropbox.createDeliveryRecord('text-meeting-cancelled-' + pending.code, {owner});
    }
    const nextHistory = [...history, {role:'user',content:text,receivedAt}, {role:'assistant',content:reply}].slice(-12);
    if (!code) { this.save(key, {owner, status:'changed', history:nextHistory}); return reply; }
    const stored = await this.dropbox.readDeliveryRecord('text-meeting-proposal-' + code);
    if (!stored || stored.owner !== owner || stored.requestKey !== requestKey || stored.reply !== reply) throw Error('Unverified email meeting proposal.');
    this.save(key, {owner, code, sourceRequestKey:requestKey, status:'prepared', presentedAt:null, history:nextHistory});
    return plain(reply);
  }
  recordReply({owner, conversationId, requestKey}) {
    const key = this.key(owner, conversationId), pending = this.pending(key);
    if (pending?.status === 'prepared' && pending.sourceRequestKey === requestKey) {
      this.save(key, {...pending, status:'pending', presentedAt:this.now().toISOString()});
    }
  }
}
