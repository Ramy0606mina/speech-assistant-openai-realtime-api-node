import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

export class DurableTaskState {
  constructor(filePath) {
    this.filePath = filePath;
    this.jobs = new Map();
    this.keys = new Map();
    if (!filePath) return;
    try {
      const data = JSON.parse(readFileSync(filePath, 'utf8'));
      this.jobs = new Map(data.jobs);
      this.keys = new Map(data.keys);
      for (const job of this.jobs.values()) {
        if (job.status === 'running') {
          job.status = 'interrupted-review';
          job.error = 'Processing interrupted by restart; inspect before retrying.';
        }
      }
      this.save();
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  save() {
    if (!this.filePath) return;
    mkdirSync(dirname(this.filePath), { recursive: true });
    const temp = this.filePath + '.tmp';
    writeFileSync(temp, JSON.stringify({ jobs: [...this.jobs], keys: [...this.keys] }));
    renameSync(temp, this.filePath);
  }
}

