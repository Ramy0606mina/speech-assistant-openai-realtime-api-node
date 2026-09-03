# London Assistant — Consolidated Production Hardening Candidate

## What this build adds

- Preserves the current voice, SMS/WhatsApp, Microsoft Graph, calendar, Action Register, Daily Brief, and Task Inbox baseline.
- Explicit sender identity for new emails: `from my email` / default executive sender = `ramy.mina@minaco.ca`; `from London` = `london@minaco.ca`.
- Dropbox read-only tools hard-scoped in code to `/LONDON - ACCESS` with path-traversal rejection.
- Messaging conversation state can persist to a JSON state file when `LONDON_STATE_FILE` points to a persistent disk.
- `/health` lightweight status endpoint.
- `/health/deep` protected diagnostic endpoint for Microsoft Graph, Twilio, and Dropbox connectivity.
- Existing Realtime watchdog/reconnect/stall protection retained.
- Existing short progress acknowledgements and concise executive-assistant prompt retained.

## Important Dropbox distinction

The Dropbox account connected to ChatGPT is not automatically available to the Render-hosted London application. For London's deployed runtime, create/authorize Dropbox API credentials for the dedicated `london@minaco.ca` Dropbox account and set `DROPBOX_ACCESS_TOKEN` in Render. Do not paste the token into chat or source code.

The code additionally enforces `DROPBOX_ROOT_PATH=/LONDON - ACCESS`. Keep the dedicated London Dropbox account limited to the shared folder as already configured.

## Render deployment

1. Do not replace the live service first. Create a staging service/branch if possible.
2. Replace `index.js` and `package.json` with this package.
3. Preserve all existing environment variables.
4. Add `HEALTH_SECRET` (strong random value).
5. Add Dropbox runtime credentials only after the Dropbox API app is authorized.
6. For restart-persistent SMS/WhatsApp context, mount a Render persistent disk at `/var/data` and set `LONDON_STATE_FILE=/var/data/london-state.json`. Without a persistent disk, use `/tmp/london-state.json` and understand that context resets on redeploy/restart.
7. Deploy and run the acceptance tests below before changing Twilio/Power Automate production webhooks.

## Security

- Do not commit `.env` or credentials to GitHub.
- `/health/deep` requires header `x-london-health-secret` equal to `HEALTH_SECRET`.
- Dropbox paths outside `/LONDON - ACCESS` are rejected.
- Dropbox tools in this build are read-only.
- Existing approval gates for consequential email/calendar actions are retained.

## Billing / subscription verification

This source code cannot prove that Render, OpenAI, Twilio, Microsoft, GitHub, or Dropbox billing will remain paid. Before production cutover, verify each provider's billing/credit/autorecharge dashboard. In particular, OpenAI API credit exhaustion previously caused Realtime calls to fail, so credit/auto-recharge should be checked separately.
