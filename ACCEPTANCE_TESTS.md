# Acceptance tests

Run on staging before production cutover.

1. `npm test` — must pass syntax validation.
2. GET `/health` — must return `ok: true`.
3. GET `/health/deep` with the correct `x-london-health-secret` — Microsoft/Twilio should be green; Dropbox should be green only after its runtime token is configured.
4. Voice: call from Ramy's authorized number; London should answer promptly and allow interruption without dropping the call.
5. Voice email identity: say "Email [known contact] from my email..."; draft readback must say sender `ramy.mina@minaco.ca`. Do not send unless Ramy explicitly says "Send it".
6. Voice London identity: say "Email [known contact] from London..."; draft readback must say sender `london@minaco.ca`.
7. Inbox: ask for latest email; then ask London to read the selected message; verify live content, no guessing.
8. Calendar: ask for three openings next week; verify returned times against Outlook.
9. SMS/WhatsApp context: send a substantive message, then "put this in tomorrow's to-do list"; verify London resolves the referent and creates the correct Action Register item.
10. Restart test: with a persistent disk configured, redeploy/restart staging and verify recent SMS/WhatsApp context is retained.
11. Dropbox: ask London to list/search `/LONDON - ACCESS`; try an outside path such as `/OtherFolder`; outside access must be rejected.
12. Task Inbox: send a harmless internal test task from Ramy's Minaco mailbox to London and verify acknowledgement, analysis, Action Register update, and completion email.
13. Daily Brief: invoke the existing scheduled endpoint manually with its secret and verify the brief uses current live data.
14. Failure test: temporarily use an invalid Dropbox token in staging; `/health/deep` should show Dropbox failed without taking down voice/email/calendar.

Do not cut over production until the tests relevant to configured integrations pass.
