# London channel expansion

Local implementation for the 20 channel entries in the owner's red-marked capability sheet. Deployment is tracked separately; this document does not claim these changes are live.

| Feature | Added or verified channels | Behavior |
|---|---|---|
| Actual Outlook draft creation | Owner email | Saves an actual draft, resolves recipient names, verifies saved draft, records current draft for follow-up revisions. |
| Existing draft body revision | Phone, owner email | Reuses verified same-ID body-only update and readback; preserves recipients, subject and reply thread. |
| Personal Outlook reminders | SMS, phone | Direct owner request authorizes save; clarify missing purpose/time, no extra confirmation, no attendees. |
| In-person and Teams meetings | SMS, phone, owner email | In-person proposals include owner-supplied location. Preserve existing meeting-confirmation behavior. SMS uses plain confirm; email uses proposal code; phone uses spoken confirmation. |
| Existing meeting cancellation | SMS, owner email | Show subject/time/organizer/attendees. Require separate exact CONFIRM CANCEL code. Re-read event and reject changed, expired, ambiguous or non-organizer targets. |
| Action Register creation | SMS, phone | Exact title/date, default due-date reminder unless opted out, durable attempt protection. |
| Task status update | SMS | Read/select existing action, update status, preserve due date; durable attempt protection. |
| Dropbox document analysis | SMS, phone | Search/list exact file first. Download actual bounded content and analyze through task model with no action tools. Retain source/limitations. |
| Dropbox rename | SMS, phone | Verify selected file identity; preserve folder/extension; no automatic alternate names or overwrite; durable attempt protection. |
| Saved PDF, Word and Excel reports | SMS, phone, owner email | Reuse actual report renderers and Dropbox London Work saving. Verify requested output paths; retain owner-email attachments. |

Ordinary email sending remains disabled in the new tools. Existing owner-only result emails and confirmed calendar invitations/cancellations retain their established behavior.

SMS routes requested actions to bounded tools and returns verified receipts. Short detail replies can continue a clarification; unrelated questions do not enable old write requests. Model-only completion text cannot substitute for a verified action.

No new dependencies or environment variables. Phone document analysis uses the existing task OpenAI client. Providers still require their established credentials and permissions. Tests use synthetic messages and fake providers; no real personal data or third-party communications are used.

Before deployment, use the normal approved PR/CI workflow. The public-repository publication gate must be resolved explicitly; do not use an alternative upload route.
