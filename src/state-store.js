import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

export class StateStore {
  constructor(filePath) {
    this.filePath = filePath;
    this.state = { processedMessages: {}, lastPollAt: null };
    this.load();
  }

  load() {
    if (!this.filePath) return this.state;
    try {
      const parsed = JSON.parse(readFileSync(this.filePath, 'utf8'));
      this.state = {
        processedMessages: parsed?.processedMessages && typeof parsed.processedMessages === 'object' ? parsed.processedMessages : {},
        lastPollAt: parsed?.lastPollAt || null,
      };
    } catch (error) {
      if (error?.code !== 'ENOENT') throw error;
    }
    return this.state;
  }

  hasMessage(key) {
    return Boolean(key && Object.hasOwn(this.state.processedMessages, key));
  }

  markMessage(key, metadata = {}) {
    if (!key) return;
    Object.defineProperty(this.state.processedMessages, key, { value: { processedAt: new Date().toISOString(), ...metadata }, enumerable: true, configurable: true, writable: true });
    // Retain delivery records; pruning can resend old messages still in the inbox.
    this.save();
  }

  markPoll() {
    this.state.lastPollAt = new Date().toISOString();
    this.save();
  }

  save() {
    if (!this.filePath) return;
    mkdirSync(dirname(this.filePath), { recursive: true });
    const temp = `${this.filePath}.tmp`;
    writeFileSync(temp, JSON.stringify(this.state, null, 2));
    renameSync(temp, this.filePath);
  }
}
