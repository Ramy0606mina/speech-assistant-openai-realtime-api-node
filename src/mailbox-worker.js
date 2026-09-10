// Inbox visibility remains available during a ledger outage. Never execute or
// mark messages processed until the original durable ledger is available.
export class MailboxWorker {
  constructor({ graph, london, guard, limit = 10, logger = console }) {
    Object.assign(this, { graph, london, guard, limit, logger });
    this.running = false;
    this.guardReady = false;
    this.status = { state: 'starting', lastCheckedAt: null };
  }

  async poll() {
    if (this.running) return { ok: false, busy: true };
    this.running = true;
    try {
      if (!this.guardReady) {
        try { await this.guard.initialize(); this.guardReady = true; }
        catch {
          const messages = await this.graph.listLondonInbox(this.limit);
          this.status = { state: 'blocked', reason: 'delivery-ledger-unavailable',
            lastCheckedAt: new Date().toISOString(), checked: messages.length };
          return { ok: false, ...this.status };
        }
      }
      const result = await this.london.pollOnce(this.limit);
      const failed = result.results.filter(item => item.processed === false).length;
      this.status = { state: failed ? 'degraded' : 'ready',
        lastCheckedAt: new Date().toISOString(), checked: result.checked, failed };
      return { ok: !failed, ...this.status };
    } catch {
      this.status = { state: 'blocked', reason: 'mailbox-poll-failed', lastCheckedAt: new Date().toISOString() };
      return { ok: false, ...this.status };
    } finally { this.running = false; }
  }
}
