Production target: London AI, main baseline ab832aa10e55e35ba2e1f8433d678db27eaa94c5.

Task Inbox now requests the actual owner reply in the report rather than treating it as a draft. Other recipients remain drafts; the job's outbound recipient stays fixed to RAMY_MINACO_EMAIL. It sends one completion rather than a separate acknowledgement plus completion. Action Register calls were removed from this workflow. Voice, calendar, and the existing attachment pipeline remain.

SMS intake returns empty TwiML and the outbound messaging helper refuses sends. The WhatsApp route was removed. Some shared legacy parser helpers remain for voice compatibility.

Exact Graph message IDs are required at intake and used as duplicate keys. A failed lookup no longer substitutes another email by subject. Configure LONDON_TASK_STATE_FILE on durable storage to persist jobs. Queued jobs resume after restart; running jobs and uncertain completion sends require review instead of blind retries. Without that setting, the existing in-memory behavior remains and restart durability is not provided. No paid storage is provisioned by this change. The verified production service currently has no attached disk.

Health reports the deployed revision, disabled messaging channels, configured state storage, and tasks requiring review. Graph acceptance does not prove email delivery. Model and email API tests use mocks; live acceptance remains a separate check.

Rollback uses the prior Render deployment. The patch does not migrate the application to the staging architecture or change hosting plans.
