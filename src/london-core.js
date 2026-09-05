function emailAddress(message) {
  return String(message?.from?.emailAddress?.address || '').trim().toLowerCase();
}

function messageKey(message) {
  return String(message?.internetMessageId || message?.id || '').trim();
}

function completionSubject(subject) {
  return `LONDON — Task Complete | ${String(subject || '(no subject)').trim() || '(no subject)'}`;
}

export class LondonCore {
  constructor({ graph, openai, dropbox, state, logger = console }) {
    this.graph = graph;
    this.openai = openai;
    this.dropbox = dropbox;
    this.state = state;
    this.logger = logger;
    this.inFlight = new Set();
  }

  async processMessage(summary) {
    const key = messageKey(summary);
    if (!key) return { skipped: true, reason: 'missing-message-id' };
    if (this.state.hasMessage(key)) return { skipped: true, reason: 'duplicate', key };

    if (this.inFlight.has(key)) return { skipped: true, reason: 'in-flight', key };
    this.inFlight.add(key);
    try {
      return await this.processClaimedMessage(summary, key);
    } finally { this.inFlight.delete(key); }
  }

  async processClaimedMessage(summary, key) {
    const full = await this.graph.getLondonMessage(summary.id);
    const sender = emailAddress(full);
    const principal = this.graph.principalMailbox;
    const london = this.graph.readMailbox;

    if (sender && sender === london) {
      this.state.markMessage(key, { sender, result: 'self-message' });
      return { skipped: true, reason: 'self-message', key };
    }

    let result;
    const actualSender = String(full.sender?.emailAddress?.address || sender).trim().toLowerCase();
    if (principal && sender === principal && actualSender === principal) {
      const attachments = full.hasAttachments ? await this.graph.getLondonAttachments(summary.id) : [];
      const analysis = await this.openai.analyzeDelegatedEmail(full, attachments, { dropbox: this.dropbox });
      const text = String(analysis.text || '').trim();
      if (!text) throw new Error('London produced an empty delegated-task result.');

      // Persist before dispatch: an interrupted/ambiguous send must not be retried blindly.
      this.state.markMessage(key, { sender, result: 'delivery-pending-review' });
      await this.graph.sendMail({
        to: principal,
        subject: completionSubject(full.subject),
        body: text,
      });

      result = { type: 'delegated-task', sender, analysis: text, completionSent: true };
    } else {
      const classification = await this.openai.classifyInboundEmail(full);
      result = { type: 'inbound-email', sender, classification: classification.text.trim().toUpperCase() };
    }

    this.state.markMessage(key, { sender, result: result.type });
    return { processed: true, key, ...result };
  }

  async pollOnce(limit = 10) {
    const messages = await this.graph.listLondonInbox(limit);
    const results = [];
    for (const message of [...messages].reverse()) {
      try {
        results.push(await this.processMessage(message));
      } catch (error) {
        this.logger.error?.('London message processing failed', { messageId: message?.id, error: error?.message });
        results.push({ processed: false, messageId: message?.id, error: error?.message || String(error) });
      }
    }
    this.state.markPoll();
    return { checked: messages.length, results };
  }
}
