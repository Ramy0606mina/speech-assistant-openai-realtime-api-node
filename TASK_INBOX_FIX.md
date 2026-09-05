# Task Inbox correction — 2026-09-05

Based on production-candidate commit 36a2c4e88701abcf3315094437edf700a4c3c2de.

The existing code already sent the model output to the principal, but asked the model for an execution brief. It now requests the actual task result, including exact phrases and receipt confirmations. Automatic Graph sends are restricted in code to the configured principal, including all To and Cc recipients. Third-party requests remain labelled drafts within the owner's completion email.

Document attachments are fetched from Graph and included in the model request. Inline signatures are omitted; unsupported files are explicitly identified. Supported file extensions: PDF, DOC/DOCX, XLS/XLSX, PPT/PPTX, TXT, CSV, MD, RTF. Limits: 20 non-inline attachments and 20 MB of document bytes. Missing attachment data fails the task rather than claiming analysis.

Duplicate protection covers concurrent calls and persisted processed messages. A delivery-pending-review record is written before Graph submission. If sending fails or the process stops during submission, inspect London's Sent Items and the record before manually clearing it for a retry. Graph's sendMail acceptance is not proof of recipient delivery, and a timeout cannot safely be automatically retried. Analysis failures remain retryable.

Deployment requirements:

- Confirm the Render service is tracking the intended production-candidate branch before deploying.
- Use one service instance and a persistent disk; set LONDON_STATE_FILE to a file on that disk (for example /var/data/london-lean-state.json). The existing /tmp default is not durable across deployments. This patch does not provision a paid disk or claim durability without one.
- Preserve existing credentials, owner mailbox, voice, and calendar settings.
- Health now exposes the Render commit, pending delivery review count, SMS paused, and WhatsApp removed. It does not expose credentials.
- The current baseline has no SMS/WhatsApp handlers or Action Register dependency; voice and calendar read functionality remain present.

Verification:

- Run npm test for existing and new regression tests.
- After deployment, send a fresh owner email to london@minaco.ca with subject "London Test 1 - Basic Task", asking for receipt confirmation, the same subject, and the exact phrase "London Task Inbox is operational".
- Confirm one completion arrives in ramy.mina@minaco.ca, then wait through another poll and confirm no duplicate.
- Repeat with a small PDF containing a unique fact and verify the answer cites that fact.
- Do not equate mocked API tests with live model accuracy, email delivery, or a voice call.

Remaining operational limits: polling still reads the latest configured inbox batch; a backlog larger than the batch needs separate catch-up work. Owner recognition uses Graph from/sender addresses and relies on the tenant's normal anti-spoofing controls. No new external-action or calendar-write tools are introduced.
