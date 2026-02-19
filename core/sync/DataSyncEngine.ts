import {
  DataReference,
  DataChange,
  DataSource,
  DocumentId,
  ReferenceId,
} from '../types';
import { eventBus, Events } from '../eventBus';
import { DataLinkManager } from '../editors/DataLinkManager';
import type { Conflict as SyncConflict, ResolveStrategy } from './ConflictResolver';
import type { Conflict } from './ConflictResolver';

// Re-export from ConflictResolver to avoid duplicates
export type {
  ResolveStrategy,
  Conflict as SyncConflict,
  ConflictType,
  ConflictResolution,
  MergeResult,
} from './ConflictResolver';

/**
 * Data Sync Engine
 *
 * Handles manual and automatic synchronization of data references
 * between Word documents and Excel data sources.
 */
export interface DataSyncEngine {
  // Manual Refresh
  refreshReference(id: ReferenceId): Promise<void>;
  refreshDocument(docId: DocumentId): Promise<void>;
  refreshAll(): Promise<void>;
  /** Refresh all references that use the given data source (selective refresh by source). */
  refreshBySource(source: DataSource): Promise<void>;

  // Auto Sync
  enableAutoSync(enabled: boolean): void;
  setSyncInterval(interval: number): void;
  isAutoSyncEnabled(): boolean;

  // Conflict Detection
  detectConflicts(): SyncConflict[];
  resolveConflict(conflict: SyncConflict, strategy: ResolveStrategy): void;

  // Event Handlers
  onDataChange(callback: (changes: DataChange[]) => void): void;
  offDataChange(callback: (changes: DataChange[]) => void): void;

  // Cleanup
  destroy(): void;
}

/**
 * Default implementation of DataSyncEngine
 */
export class DefaultDataSyncEngine implements DataSyncEngine {
  private autoSyncEnabled = false;
  private syncInterval = 60000; // 1 minute default
  private syncTimer: ReturnType<typeof setInterval> | null = null;
  private dataChangeCallbacks: Set<(changes: DataChange[]) => void> = new Set();

  constructor(
    private dataLinkManager: DataLinkManager,
    private dataSourceResolver: (ref: DataReference) => Promise<any>
  ) {}

  /**
   * Refresh a single reference
   */
  async refreshReference(id: ReferenceId): Promise<void> {
    const ref = this.dataLinkManager.getReference(id);
    if (!ref) {
      console.warn(`Reference ${id} not found`);
      return;
    }

    const oldValue = ref.display.value;
    const newValue = await this.dataSourceResolver(ref);

    // Update the reference with new value
    this.dataLinkManager.updateReference(id, {
      display: {
        ...ref.display,
        value: newValue,
      },
      state: 'active',
      history: [
        ...ref.history,
        {
          timestamp: Date.now(),
          action: 'updated',
          value: newValue,
          userId: 'system',
        },
      ],
    });

    // Emit data change event if value changed
    if (JSON.stringify(oldValue) !== JSON.stringify(newValue)) {
      const change: DataChange = {
        referenceId: id,
        oldValue,
        newValue,
        timestamp: Date.now(),
        source: 'manual',
      };

      eventBus.emit(Events.DataChanged, [change]);
      this.notifyDataChange([change]);
    }
  }

  /**
   * Refresh all references in a document
   */
  async refreshDocument(docId: DocumentId): Promise<void> {
    const refIds = this.dataLinkManager.findReferencesByDocument(docId);
    const refs = refIds
      .map((id) => this.dataLinkManager.getReference(id))
      .filter((r): r is DataReference => !!r);

    if (refs.length === 0) return;

    // Batch resolve for better performance
    const values = await this.dataLinkManager.resolveReferenceBatch(refs);
    const changes: DataChange[] = [];

    for (const ref of refs) {
      const oldValue = ref.display.value;
      const newValue = values.get(ref.id);

      if (newValue !== undefined) {
        this.dataLinkManager.updateReference(ref.id, {
          display: {
            ...ref.display,
            value: newValue,
          },
          state: 'active',
        });

        if (JSON.stringify(oldValue) !== JSON.stringify(newValue)) {
          changes.push({
            referenceId: ref.id,
            oldValue,
            newValue,
            timestamp: Date.now(),
            source: 'manual',
          });
        }
      }
    }

    if (changes.length > 0) {
      eventBus.emit(Events.DataChanged, changes);
      this.notifyDataChange(changes);
    }
  }

  /**
   * Refresh all references
   */
  async refreshAll(): Promise<void> {
    const refs = this.dataLinkManager.getAllReferences();

    if (refs.length === 0) return;

    // Batch resolve for better performance
    const values = await this.dataLinkManager.resolveReferenceBatch(refs);
    const changes: DataChange[] = [];

    for (const ref of refs) {
      const oldValue = ref.display.value;
      const newValue = values.get(ref.id);

      if (newValue !== undefined) {
        this.dataLinkManager.updateReference(ref.id, {
          display: {
            ...ref.display,
            value: newValue,
          },
          state: 'active',
        });

        if (JSON.stringify(oldValue) !== JSON.stringify(newValue)) {
          changes.push({
            referenceId: ref.id,
            oldValue,
            newValue,
            timestamp: Date.now(),
            source: 'manual',
          });
        }
      }
    }

    if (changes.length > 0) {
      eventBus.emit(Events.DataChanged, changes);
      this.notifyDataChange(changes);
    }
  }

  /**
   * Refresh all references that use the given data source (selective refresh by source).
   */
  async refreshBySource(source: DataSource): Promise<void> {
    const refIds = this.dataLinkManager.findReferencesBySource(source);
    const refs = refIds
      .map((id) => this.dataLinkManager.getReference(id))
      .filter((r): r is DataReference => !!r);

    if (refs.length === 0) return;

    const values = await this.dataLinkManager.resolveReferenceBatch(refs);
    const changes: DataChange[] = [];

    for (const ref of refs) {
      const oldValue = ref.display.value;
      const newValue = values.get(ref.id);

      if (newValue !== undefined) {
        this.dataLinkManager.updateReference(ref.id, {
          display: {
            ...ref.display,
            value: newValue,
          },
          state: 'active',
        });

        if (JSON.stringify(oldValue) !== JSON.stringify(newValue)) {
          changes.push({
            referenceId: ref.id,
            oldValue,
            newValue,
            timestamp: Date.now(),
            source: 'manual',
          });
        }
      }
    }

    if (changes.length > 0) {
      eventBus.emit(Events.DataChanged, changes);
      this.notifyDataChange(changes);
    }
  }

  /**
   * Enable or disable automatic synchronization
   */
  enableAutoSync(enabled: boolean): void {
    this.autoSyncEnabled = enabled;

    if (enabled) {
      this.startAutoSync();
    } else {
      this.stopAutoSync();
    }
  }

  /**
   * Set the auto-sync interval in milliseconds
   */
  setSyncInterval(interval: number): void {
    this.syncInterval = interval;

    if (this.autoSyncEnabled) {
      this.stopAutoSync();
      this.startAutoSync();
    }
  }

  /**
   * Check if auto-sync is enabled
   */
  isAutoSyncEnabled(): boolean {
    return this.autoSyncEnabled;
  }

  /**
   * Start auto-sync timer
   */
  private startAutoSync(): void {
    if (this.syncTimer) return;

    this.syncTimer = setInterval(() => {
      this.refreshAll().catch((error) => {
        console.error('Auto-sync error:', error);
      });
    }, this.syncInterval);
  }

  /**
   * Stop auto-sync timer
   */
  private stopAutoSync(): void {
    if (this.syncTimer) {
      clearInterval(this.syncTimer);
      this.syncTimer = null;
    }
  }

  /**
   * Detect conflicts in references
   *
   * In a real implementation, this would check against
   * concurrent edits or version mismatches.
   */
  detectConflicts(): Conflict[] {
    const conflicts: Conflict[] = [];
    const refs = this.dataLinkManager.getAllReferences();

    for (const ref of refs) {
      if (ref.state === 'conflict') {
        conflicts.push({
          id: crypto.randomUUID(),
          type: 'data-inconsistency',
          referenceId: ref.id,
          reference: ref,
          localValue: ref.display.value,
          remoteValue: null, // Would be populated from conflict detection
          localTimestamp: Date.now(),
          remoteTimestamp: Date.now(),
          resolved: false,
        });
      }
    }

    return conflicts;
  }

  /**
   * Resolve a conflict with the specified strategy
   */
  resolveConflict(conflict: Conflict, strategy: ResolveStrategy): void {
    const ref = this.dataLinkManager.getReference(conflict.referenceId);
    if (!ref) return;

    switch (strategy) {
      case 'keep-local':
        // Keep local value, mark as active
        this.dataLinkManager.updateReference(conflict.referenceId, {
          state: 'active',
          display: {
            ...ref.display,
            value: conflict.localValue,
          },
        });
        break;

      case 'keep-remote':
        // Use remote value
        this.dataLinkManager.updateReference(conflict.referenceId, {
          state: 'active',
          display: {
            ...ref.display,
            value: conflict.remoteValue,
          },
        });
        break;

      case 'merge':
        // Implement merge strategy (e.g., timestamp-based)
        // For now, default to keeping remote
        this.dataLinkManager.updateReference(conflict.referenceId, {
          state: 'active',
          display: {
            ...ref.display,
            value: conflict.remoteValue,
          },
        });
        break;
    }

    // Emit event for conflict resolution
    eventBus.emit(Events.ReferenceUpdated, {
      referenceId: conflict.referenceId,
      strategy,
    });
  }

  /**
   * Register a data change listener
   */
  onDataChange(callback: (changes: DataChange[]) => void): void {
    this.dataChangeCallbacks.add(callback);
  }

  /**
   * Unregister a data change callback (mirrors onDataChange).
   */
  offDataChange(callback: (changes: DataChange[]) => void): void {
    this.dataChangeCallbacks.delete(callback);
  }

  /**
   * Notify all registered data change listeners
   */
  private notifyDataChange(changes: DataChange[]): void {
    this.dataChangeCallbacks.forEach((callback) => {
      try {
        callback(changes);
      } catch (error) {
        console.error('Error in data change callback:', error);
      }
    });
  }

  /**
   * Cleanup resources
   */
  destroy(): void {
    this.stopAutoSync();
    this.dataChangeCallbacks.clear();
  }
}

/**
 * Factory function to create a DataSyncEngine instance
 */
export function createDataSyncEngine(
  dataLinkManager: DataLinkManager,
  dataSourceResolver?: (ref: DataReference) => Promise<any>
): DataSyncEngine {
  const resolver =
    dataSourceResolver ??
    ((ref: DataReference) => dataLinkManager.resolveReference(ref));

  return new DefaultDataSyncEngine(dataLinkManager, resolver);
}
