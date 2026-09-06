import { createHash } from 'node:crypto';

// Durable, fail-closed at-most-once dispatch. A send timeout requires review,
// not another automatic send. No document contents are stored in the ledger.
export class DeliveryGuard {
  constructor(store, mailbox) {
    this.store = store;
    this.prefix = createHash('sha256').update(String(mailbox).trim().toLowerCase()).digest('hex');
  }
  record(key) { return `${this.prefix}-${createHash('sha256').update(key).digest('hex')}`; }
  async initialize() {
    const key = `${this.prefix}-cutover`;
    let record = await this.store.readDeliveryRecord(key);
    if (!record) {
      await this.store.createDeliveryRecord(key, { notBefore: new Date().toISOString(), result: 'migration-boundary' });
      record = await this.store.readDeliveryRecord(key);
    }
    if (!record || !Number.isFinite(Date.parse(record.notBefore))) throw new Error('Delivery ledger boundary is missing or invalid.');
    this.notBefore = Date.parse(record.notBefore);
  }
  async check(key, receivedDateTime) {
    if (!Number.isFinite(this.notBefore)) throw new Error('Delivery ledger is not initialized.');
    const received = Date.parse(receivedDateTime);
    if (!Number.isFinite(received) || received < this.notBefore) return 'historical-review';
    if (await this.store.readDeliveryRecord(this.record(key))) return 'durable-duplicate';
    return null;
  }
  async claim(key) {
    return this.store.createDeliveryRecord(this.record(key), { result: 'delivery-pending-review', claimedAt: new Date().toISOString() });
  }
  async complete(key, reportPath) {
    // Separate receipt preserves the immutable dispatch claim.
    await this.store.createDeliveryRecord(`${this.record(key)}-sent`, { result: 'sent', sentAt: new Date().toISOString(), reportPath: reportPath || null });
  }
}
