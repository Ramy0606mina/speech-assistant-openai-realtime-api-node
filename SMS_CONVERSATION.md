# Owner SMS conversations

London receives texts from the existing RAMY_PHONE_NUMBER to TWILIO_PHONE_NUMBER and replies only to RAMY_PHONE_NUMBER. General conversation and email, calendar, action-register and Dropbox lookups are supported. Owner texts can create personal Outlook reminders, prepare confirmed Teams invitations, and save or revise Outlook email drafts. Photos and documents are acknowledged with a request to email them for analysis.

## Outlook email workflow

- Ask London to draft or prepare an email, or reply to an existing email. Names are resolved through the existing contact tool; replies use Microsoft's native reply draft to preserve the thread.
- London saves the actual message in the principal's Outlook Drafts folder by default, reads the saved ID back, and returns its recipient and subject. A drafting request never sends the email.
- Reply with a missing recipient or `save it` to continue an unfinished drafting request. After saving, request wording changes to update the same draft body in place. Recipients and subject are preserved by this revision tool; edit those in Outlook if needed.
- Email handling by SMS is draft-only. London never sends an email or asks for permission to send. The saved receipt only confirms the mailbox, recipient and subject. A `send it` text receives a draft-only response, even in older conversations that offered sending.
- `Keep it in my drafts` leaves the message in Outlook. Review or edit it there whenever needed; the provider retains it even if London's conversation state is lost. `Review the draft` refreshes the saved reference.
- New drafts, reply drafts and revisions use plain text, normalize line endings and remove leading/trailing whitespace. Blank lines between paragraphs are preserved. No HTML paragraphs, margins or leading blank lines are generated. The mail application's own composer padding is outside the email body.

The Microsoft application credentials used by `voiceRequest` need `Mail.ReadWrite` to create/update drafts. SMS has no email-send method or tool. This code does not change Microsoft permissions used by other services. Reference: [update a draft](https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0).

The production source branch is `london-production-candidate-2026-09-02`, starting with `node lean-server.js`. The legacy `main` branch is a different server and must not replace this build. The authenticated Render dashboard confirms the active service is `srv-da6r7i61egvs73bsfpd0` at `https://london-ai.onrender.com`, with draft-saving commit `8cdd3f6` live during the September 12 repair. The previously documented `london-ai-pr-1.onrender.com` returned a suspended response but is not the active production URL. Do not resume or deploy that older service. The SMS draft workflow supports saving and revising drafts only; it has no sending action.

The default URL check in code still names the older https://london-ai-pr-1.onrender.com deployment. The active https://london-ai.onrender.com service must explicitly set LONDON_SMS_CONVERSATION_ENABLED=true. Verify the existing setting before changing it. Set it to false to disable processing and automatic webhook setup. Other deployments and PR previews must not process production texts.

At startup the service verifies the exact SMS-capable Twilio number, configures only its SMS URL/method and fallback to /incoming-sms, and reads the settings back. Voice configuration is untouched. A configured TwiML application is left alone and reported as requiring review. Messaging Services can override the number webhook; the Twilio inbox poll remains the source of truth. Review a Messaging Service separately if it still generates an unwanted automatic response.

The signed POST webhook responds immediately with empty TwiML. A serialized worker retrieves the original messages from Twilio every 15 seconds and when notified by the webhook, so model latency and webhook retries cannot lose the request. A durable Dropbox cutover excludes texts sent before initial activation. Pagination is validated and messages are processed chronologically. Durable dispatch claims prevent duplicate SMS replies after restarts; an uncertain send is held for review instead of retried. This favors avoiding duplicate replies over automatic retransmission of an uncertain send.

The last six exchanges are kept in the existing local state file for up to 24 hours of conversational use. A persistent LONDON_STATE_FILE retains them across restarts; an ephemeral disk can lose context on redeployment. Twilio retains the original messages. STOP/START/HELP are left to Twilio; pending older requests are suppressed by a later stop/start control, and a local opt-out prevents replies until START. Existing urgent-alert behavior is unchanged and still subject to Twilio opt-out enforcement.

/health includes enabled/configured/ready, webhook verification, lastCheckedAt, lastOutcome, and lastReplyStatus. Accepted/queued means Twilio accepted the reply, not that the handset received it. Only delivered confirms Twilio delivery. No phone numbers or message contents are exposed by health.

Validation: send a new text from the configured owner to London's existing number after the deployment is ready. Verify an inbound request and a delivered reply, then send a follow-up to check conversational context. No live delivery should be claimed solely from automated tests or health readiness.

References: https://www.twilio.com/docs/messaging/api/message-resource, https://www.twilio.com/docs/messaging/guides/webhook-request, https://www.twilio.com/docs/phone-numbers/api/incomingphonenumber-resource
