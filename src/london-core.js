function emailAddress(message) {
  return String(message?.from?.emailAddress?.address || '').trim().toLowerCase();
}

function messageKey(message) {
  return String(message?.internetMessageId || message?.id || '').trim();
}

export class LondonCore {
  constructor({ graph, openai, dropbox, state, logger = console }) {
    this.graph = graph;
    this.openai = openai;
    this.dropbox = dropbox;
    this.state = state;
    this.logger = logger;
  }

  async processMessage(summary) {
    const key = messageKey(summary);
    if (!key) return { skipped: true, reason: 'missing-message-id' };
    if (this.state.hasMessage(key)) return { skipped: true, reason: 'duplicate', key };

    const full = await this.graph.getLondonMessage(summary.id);
    const sender = emailAddress(full);
    const principal = this.graph.principalMailbox;
    const london = this.graph.readMailbox;

    if (sender && sender === london) {
      this.state.markMessage(key, { sender, result: 'self-message' });
      return { skipped: true, reason: 'self-message', key };
    }

    let result;
    if (principal && sender === principal) {
      const analysis = await this.openai.analyzeDelegatedEmail(full);
      result = { type: 'delegated-task', sender, analysis: analysis.text };
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
