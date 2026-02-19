/**
 * Update Queue for Data Synchronization
 *
 * Manages pending updates with batching and deduplication
 */

export interface QueuedUpdate {
  id: string;
  referenceId: string;
  priority: 'high' | 'normal' | 'low';
  timestamp: number;
  retryCount: number;
  updateFn: () => Promise<void>;
}

export interface UpdateQueueConfig {
  maxBatchSize: number;
  batchSize: number;
  maxRetries: number;
  retryDelay: number;
  flushInterval: number;
}

const DEFAULT_CONFIG: UpdateQueueConfig = {
  maxBatchSize: 100,
  batchSize: 10,
  maxRetries: 3,
  retryDelay: 1000,
  flushInterval: 50, // ms
};

export class UpdateQueue {
  private queue: QueuedUpdate[] = [];
  private processing = false;
  private flushTimer: ReturnType<typeof setInterval> | null = null;
  private readonly config: UpdateQueueConfig;
  private readonly pendingUpdates = new Map<string, QueuedUpdate>();
  private readonly updateCallbacks = new Set<(updates: QueuedUpdate[]) => void>();

  constructor(config: Partial<UpdateQueueConfig> = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * Add an update to the queue
   */
  queueUpdate(
    referenceId: string,
    updateFn: () => Promise<void>,
    priority: QueuedUpdate['priority'] = 'normal'
  ): string {
    const id = `${referenceId}-${Date.now()}-${Math.random()}`;

    const update: QueuedUpdate = {
      id,
      referenceId,
      priority,
      timestamp: Date.now(),
      retryCount: 0,
      updateFn,
    };

    // Deduplicate: remove existing updates for the same reference
    this.deduplicate(referenceId);

    // Add to queue and pending map
    this.pendingUpdates.set(referenceId, update);

    // Insert based on priority
    this.insertByPriority(update);

    return id;
  }

  /**
   * Remove any pending updates for the given reference
   */
  private deduplicate(referenceId: string): void {
    const existing = this.pendingUpdates.get(referenceId);
    if (existing) {
      this.queue = this.queue.filter(u => u.id !== existing.id);
    }
  }

  /**
   * Insert update by priority (high -> normal -> low)
   */
  private insertByPriority(update: QueuedUpdate): void {
    const priorityOrder = { high: 0, normal: 1, low: 2 };
    const insertIndex = this.queue.findIndex(
      u => priorityOrder[u.priority] > priorityOrder[update.priority]
    );

    if (insertIndex === -1) {
      this.queue.push(update);
    } else {
      this.queue.splice(insertIndex, 0, update);
    }
  }

  /**
   * Start processing the queue
   */
  start(): void {
    if (this.flushTimer) return;

    this.flushTimer = setInterval(() => {
      this.flush();
    }, this.config.flushInterval);
  }

  /**
   * Stop processing the queue
   */
  stop(): void {
    if (this.flushTimer) {
      clearInterval(this.flushTimer);
      this.flushTimer = null;
    }
  }

  /**
   * Process a batch of updates
   */
  private async flush(): Promise<void> {
    if (this.processing || this.queue.length === 0) return;

    this.processing = true;

    const batchSize = Math.min(this.queue.length, this.config.batchSize);
    const batch = this.queue.splice(0, batchSize);

    try {
      // Process updates in parallel for better performance
      await Promise.all(
        batch.map(async (update) => {
          try {
            await update.updateFn();
            this.pendingUpdates.delete(update.referenceId);
          } catch (error) {
            this.handleUpdateError(update, error);
          }
        })
      );

      // Notify callbacks of completed updates
      const completed = batch.filter(u => !this.pendingUpdates.has(u.referenceId));
      if (completed.length > 0) {
        this.notifyCallbacks(completed);
      }
    } catch (error) {
      console.error('Error processing update batch:', error);
    } finally {
      this.processing = false;
    }
  }

  /**
   * Handle update errors with retry logic
   */
  private handleUpdateError(update: QueuedUpdate, error: unknown): void {
    console.error(`Update failed for ${update.referenceId}:`, error);

    if (update.retryCount < this.config.maxRetries) {
      update.retryCount++;

      // Re-queue with delay
      setTimeout(() => {
        this.queue.push(update);
        this.pendingUpdates.set(update.referenceId, update);
      }, this.config.retryDelay * update.retryCount);
    } else {
      // Max retries reached, remove from pending
      this.pendingUpdates.delete(update.referenceId);

      // Notify callbacks of failed update
      this.notifyCallbacks([update]);
    }
  }

  /**
   * Register a callback for completed updates
   */
  onUpdate(callback: (updates: QueuedUpdate[]) => void): () => void {
    this.updateCallbacks.add(callback);

    return () => {
      this.updateCallbacks.delete(callback);
    };
  }

  /**
   * Notify all registered callbacks
   */
  private notifyCallbacks(updates: QueuedUpdate[]): void {
    this.updateCallbacks.forEach((callback) => {
      try {
        callback(updates);
      } catch (error) {
        console.error('Error in update callback:', error);
      }
    });
  }

  /**
   * Get queue statistics
   */
  getStats() {
    return {
      queueSize: this.queue.length,
      pendingSize: this.pendingUpdates.size,
      processing: this.processing,
      byPriority: {
        high: this.queue.filter(u => u.priority === 'high').length,
        normal: this.queue.filter(u => u.priority === 'normal').length,
        low: this.queue.filter(u => u.priority === 'low').length,
      },
    };
  }

  /**
   * Clear all pending updates
   */
  clear(): void {
    this.queue = [];
    this.pendingUpdates.clear();
  }

  /**
   * Get all pending updates
   */
  getPendingUpdates(): QueuedUpdate[] {
    return Array.from(this.pendingUpdates.values());
  }

  /**
   * Get updates for a specific reference
   */
  getUpdatesForReference(referenceId: string): QueuedUpdate[] {
    return Array.from(this.pendingUpdates.values()).filter(
      u => u.referenceId === referenceId
    );
  }

  /**
   * Cancel updates for a specific reference
   */
  cancelUpdatesForReference(referenceId: string): number {
    const count = this.getUpdatesForReference(referenceId).length;

    this.queue = this.queue.filter(u => u.referenceId !== referenceId);
    this.pendingUpdates.delete(referenceId);

    return count;
  }

  /**
   * Destroy the queue
   */
  destroy(): void {
    this.stop();
    this.clear();
    this.updateCallbacks.clear();
  }
}

/**
 * Factory function to create an UpdateQueue instance
 */
export function createUpdateQueue(config?: Partial<UpdateQueueConfig>): UpdateQueue {
  return new UpdateQueue(config);
}
