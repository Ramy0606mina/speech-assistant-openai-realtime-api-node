import {createHash} from 'node:crypto';

const hash=value=>createHash('sha256').update(value).digest('hex');
const normalized=value=>String(value||'').trim().replace(/\s+/g,' ').toLowerCase();
export const revisionBody=value=>String(value||'').replace(/\r\n?/g,'\n').trim();
const identity=draft=>JSON.stringify({id:draft.id,subject:draft.subject,conversationId:draft.conversationId,
  from:draft.from,sender:draft.sender,replyTo:draft.replyTo,to:draft.toRecipients,cc:draft.ccRecipients,bcc:draft.bccRecipients,hasAttachments:draft.hasAttachments});
export const revisionVersion=draft=>hash(identity(draft)+JSON.stringify({changeKey:draft.changeKey,body:draft.body}));
export const draftSummary=draft=>({id:draft.id,subject:draft.subject||'(No subject)',
  recipients:(draft.toRecipients||[]).map(item=>({name:item.emailAddress?.name||'',address:item.emailAddress?.address||''}))});

export async function readDraft(graph,mailbox,id) {
  graph.voiceMailbox(mailbox); // Reject any mailbox outside the two connected identities.
  const draft=await graph.getVoiceMessage(mailbox,id);
  if(!draft || draft.id!==id || draft.isDraft!==true)throw Error('That message is no longer an available Outlook draft. Nothing was changed.');
  return draft;
}

export function matchingDrafts(drafts,selectors=[]) {
  return drafts.filter(draft=>draft.isDraft===true && selectors.every(selector=>{
    const value=normalized(selector);
    return value && [draft.subject,...(draft.toRecipients||[]).flatMap(item=>[item.emailAddress?.name,item.emailAddress?.address])]
      .some(field=>normalized(field).includes(value));
  }));
}

export async function selectDraft({graph,mailbox='principal',selectors=[],current}) {
  if(!['principal','london'].includes(mailbox))throw Error('Choose the principal or London mailbox.');
  if(!selectors.length && current?.id && current.mailbox===mailbox)return {draft:await readDraft(graph,mailbox,current.id)};
  const candidates=matchingDrafts(await graph.listEditableDrafts(mailbox),selectors);
  if(candidates.length!==1)return {candidates:candidates.map(draftSummary)};
  return {draft:await readDraft(graph,mailbox,candidates[0].id)};
}

// Shared by phone and owner email. All provider changes are body-only and target
// the exact draft read before composition; retries never repeat an uncertain write.
export async function reviseDraft({graph,dropbox,owner,sourceKey,mailbox='principal',expected,body}) {
  if(!owner || owner.toLowerCase()!==graph.principalMailbox?.toLowerCase() || !sourceKey)throw Error('Authenticated owner and source required.');
  const content=revisionBody(body);
  if(!content || content.length>20000)throw Error('The revised draft body must contain 1–20,000 characters.');
  if(!expected?.id || expected.isDraft!==true)throw Error('Read and select an existing draft before revising it.');
  const live=await readDraft(graph,mailbox,expected.id);
  if(revisionVersion(live)!==revisionVersion(expected))throw Error('The draft changed while it was being revised. Please read it again before retrying.');
  if(revisionBody(live.body?.content)===content)return live;
  const key='draft-revision-'+hash(owner.toLowerCase()+':'+sourceKey);
  if(!await dropbox.createDeliveryRecord(key,{status:'attempted',at:new Date().toISOString()}))throw Error('This draft revision was already attempted. Check Outlook before retrying.');
  const versionKey='draft-revision-version-'+hash(owner.toLowerCase()+':'+mailbox+':'+live.id+':'+revisionVersion(live));
  if(!await dropbox.createDeliveryRecord(versionKey,{status:'attempted',at:new Date().toISOString()}))throw Error('This version of the draft was already being revised. Read Outlook again before retrying.');
  await graph.updateEmailDraftBody({mailbox,id:live.id,body:content});
  const saved=await readDraft(graph,mailbox,live.id);
  if(identity(saved)!==identity(live) || revisionBody(saved.body?.content)!==content)throw Error('Outlook did not verify the complete body-only update. Check the draft before retrying.');
  await dropbox.createDeliveryRecord(key+'-done',{status:'verified',at:new Date().toISOString()});
  return saved;
}

export function revisionReceipt(graph,mailbox,draft) {
  return 'Updated the existing draft in Outlook Drafts for '+graph.voiceMailbox(mailbox)+'.\nSubject: '+(draft.subject||'(No subject)')+'\nRecipients, subject and reply thread were preserved. Nothing was sent.';
}
