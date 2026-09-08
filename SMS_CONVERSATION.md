# Owner SMS conversations

London receives texts from the existing RAMY_PHONE_NUMBER to TWILIO_PHONE_NUMBER and replies only to RAMY_PHONE_NUMBER. General conversation and read-only email, calendar, action-register and Dropbox lookups are supported. Changes to meetings, email drafts, and files continue through the existing voice/email channels. Photos and documents are acknowledged with a request to email them for analysis.

The deployed London mailbox service enables this feature by default only at https://london-ai-pr-1.onrender.com. Other deployments must explicitly set LONDON_SMS_CONVERSATION_ENABLED=true. Set it to false to disable processing and automatic webhook setup. This prevents PR previews from handling production texts.

At startup the service verifies the exact SMS-capable Twilio number, configures only its SMS URL/method and fallback to /incoming-sms, and reads the settings back. Voice configuration is untouched. A configured TwiML application is left alone and reported as requiring review. Messaging Services can override the number webhook; the Twilio inbox poll remains the source of truth. Review a Messaging Service separately if it still generates an unwanted automatic response.

The signed POST webhook responds immediately with empty TwiML. A serialized worker retrieves the original messages from Twilio every 15 seconds and when notified by the webhook, so model latency and webhook retries cannot lose the request. A durable Dropbox cutover excludes texts sent before initial activation. Pagination is validated and messages are processed chronologically. Durable dispatch claims prevent duplicate SMS replies after restarts; an uncertain send is held for review instead of retried. This favors avoiding duplicate replies over automatic retransmission of an uncertain send.

The last six exchanges are kept in the existing local state file for up to 24 hours of conversational use. A persistent LONDON_STATE_FILE retains them across restarts; an ephemeral disk can lose context on redeployment. Twilio retains the original messages. STOP/START/HELP are left to Twilio; pending older requests are suppressed by a later stop/start control, and a local opt-out prevents replies until START. Existing urgent-alert behavior is unchanged and still subject to Twilio opt-out enforcement.

/health includes enabled/configured/ready, webhook verification, lastCheckedAt, lastOutcome, and lastReplyStatus. Accepted/queued means Twilio accepted the reply, not that the handset received it. Only delivered confirms Twilio delivery. No phone numbers or message contents are exposed by health.

Validation: send a new text from the configured owner to London's existing number after the deployment is ready. Verify an inbound request and a delivered reply, then send a follow-up to check conversational context. No live delivery should be claimed solely from automated tests or health readiness.

References: https://www.twilio.com/docs/messaging/api/message-resource, https://www.twilio.com/docs/messaging/guides/webhook-request, https://www.twilio.com/docs/phone-numbers/api/incomingphonenumber-resource
