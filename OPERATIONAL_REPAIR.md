# Worker failure reporting and connector isolation

The active lean service can keep answering HTTP health checks while every task
fails. This repair makes `/health` return `live:true` plus an operational `ok`
value, mailbox outcome counts, and a sanitized OpenAI failure reason. Its HTTP
status remains 200 for Render liveness. `/ready` returns 503 while starting,
blocked, degraded, stale, or unable to use OpenAI. It returns 200 only after a
successful recent poll with no task failures and no known OpenAI failure.
Configuration flags indicate key presence, not successful operations.

OpenAI quota exhaustion pauses further Responses API attempts on the shared
client for five minutes. Failed tasks remain pending under the existing retry
policy; the repair neither buys credits nor switches keys/models. An expired
cooldown permits another attempt; only a successful request clears the failure.

If the durable ledger cannot initialize, London may list the inbox for status
but never analyze, execute, send, or mark its messages processed. It retries
the original ledger and preserves its historical boundary. There is no
ephemeral fallback ledger and no migration or historical replay.

After a durable claim and task execution, optional Dropbox report-save failure
no longer suppresses the result email. The email explicitly reports the save
as unconfirmed and uses Needs Attention. Ambiguous dispatch remains protected
by the immutable claim. General replies use Task Response, since producing
text does not prove completion of the requested action.

Validation: 152 automated tests pass, including quota cooldown/recovery, ledger
outage/recovery with historical-mail protection, degraded poll detection, and
report-save failure with durable duplicate protection across process restarts.
Live provider acceptance remains a separate check. No provider secrets, hosting
plans, phone routing, or existing ledger entries are changed by this patch.
