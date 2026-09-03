# Lean London — Single Backend Architecture

## Day 1 decision

London's core runtime is one Node.js service on Render. Power Automate is not part of the core architecture.

Core path:

`Microsoft Graph + Dropbox -> London backend -> OpenAI -> Microsoft/Dropbox actions`

Urgent SMS will be a London-backend action through Twilio, not a Power Automate flow.
Voice will connect to the same London backend in the voice phase.

## Preserved legacy build

The previous working candidate remains untouched on branch `london-production-candidate-2026-09-02` for rollback/reference.
The lean rebuild lives on `london-lean-core-2026-09-03`.

## Day 1 components

- `lean-server.js`: one Render process, health endpoints, direct mailbox polling.
- `src/microsoft-graph.js`: direct Microsoft authentication, London inbox read, full message retrieval, and send-mail action.
- `src/openai-client.js`: one OpenAI Responses API client with output parsing that tolerates reasoning items.
- `src/dropbox-client.js`: Dropbox access restricted to the configured London root.
- `src/state-store.js`: restart-safe message deduplication state.
- `src/london-core.js`: central message routing/orchestration.
- `test/lean-core.test.js`: automated regression tests.

## Security rules

- No secrets in source code.
- External senders can be classified but cannot be treated as principal-delegated tasks.
- Only the configured principal mailbox can create delegated email tasks.
- Dropbox paths remain under `DROPBOX_ROOT_PATH`.
- Manual poll/deep-health endpoints require `HEALTH_SECRET`.

## Power Automate retirement

The lean server polls `london@minaco.ca` directly through Microsoft Graph. It does not expose or require the legacy `/task-inbox` Power Automate webhook.
Legacy flows remain untouched only until lean acceptance testing is complete, then they can be disabled.
