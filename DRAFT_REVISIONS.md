# Existing Outlook draft revisions

Phone and owner email can revise the body of an existing saved Outlook draft. SMS retains its existing draft-only workflow. No channel gains an email-send action.

## Phone

Ask London to shorten, add a point to, rewrite, or translate a saved draft. London calls select_email_draft to read the current call's verified draft, or finds it using subject/recipient phrases. Multiple matches require clarification. A draft read through read_email can also become the current selection. A newly saved draft is referenced by its returned ID and must be read before revision.

The update_saved_email_draft tool changes only that selected body. It rereads the draft before writing and verifies the same ID, draft status, subject, recipients and reply conversation after writing. Reading a different received message or an ambiguous selection clears the old revision target. Call state does not persist to another call.

## Owner email

Email London an instruction such as “Shorten the Invoice review draft to Alex” or “Translate my current Outlook draft into French.” The separate revision handler runs before general delegated-task analysis. Only the authenticated principal's own subject and unquoted request authorize the revision.

An explicit subject/recipient takes precedence over the last verified email-channel target. With no current target, the mailbox must contain one uniquely matching draft; otherwise London asks for its subject or recipient. A reply identifying one of the candidates continues the original editing request. Successful email-channel context lasts up to24hours in the state file; unrelated owner requests clear it. State loss requires a fresh unique selection and does not guess the newest draft.

The default mailbox is principal. An explicit London-mailbox selection is retained for a short revision follow-up. Rewriting preserves existing facts and signature unless the owner requests a body change. The result email reports verified update or failure; it never sends the composed draft.

## Safety and limits

- Existing drafts are read and updated, never copied into a new draft as a substitute.
- The provider PATCH contains only body with plain-text content. Subject, recipients, attachments and reply conversation are not part of that PATCH.
- Recheck draft status and version before writing; verify the complete resulting body and preserved fields afterward.
- Durable request and draft-version claims prevent repeated or concurrent writes to the same source version. Uncertain writes remain held; do not automatically retry them.
- Up to1000 draft entries across20pages can be considered. Incomplete or invalid paging fails closed.
- Body limit20,000characters. No email sending, recipient editing, subject editing, or new attachment workflow is added.
- SMS continues using its established handler and Graph updateSmsDraft compatibility method.

Validation is automated with fake provider writes. No real owner draft, email or invitation is used as a test. Health exposes draftRevisions for phone, email and SMS, bodyOnly:true, sending:false; that describes deployed code, not independent real-user acceptance.
