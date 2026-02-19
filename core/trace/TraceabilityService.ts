import {
  DocumentId,
  DocumentSnapshot,
  SnapshotId,
  DataReference,
  CellId,
  ReferenceId,
  DataSource,
} from '../types';
import { eventBus, Events } from '../eventBus';

/**
 * Audit Action Types
 */
export type AuditActionType =
  | 'reference_created'
  | 'reference_updated'
  | 'reference_deleted'
  | 'reference_refreshed'
  | 'snapshot_created'
  | 'snapshot_restored'
  | 'document_saved'
  | 'data_changed';

/**
 * Audit Log Entry
 */
export interface AuditLog {
  id: string;
  timestamp: number;
  action: AuditActionType;
  userId: string;
  metadata: Record<string, any>;
}

/**
 * Snapshot Diff Result
 */
export interface Diff {
  added: string[];
  removed: string[];
  modified: DiffChange[];
}

export interface DiffChange {
  path: string;
  oldValue: any;
  newValue: any;
}

/**
 * Audit Filters
 */
export interface AuditFilters {
  startDate?: number;
  endDate?: number;
  action?: AuditActionType;
  userId?: string;
  limit?: number;
}

/**
 * Audit Action
 */
export interface AuditAction {
  action: AuditActionType;
  userId?: string;
  metadata?: Record<string, any>;
}

/**
 * Cell Change History
 */
export interface CellChange {
  timestamp: number;
  oldValue: any;
  newValue: any;
  userId: string;
}

/**
 * Traceability Service
 *
 * Provides complete data traceability capabilities including:
 * - Version snapshots
 * - Change history
 * - Dependency analysis
 * - Audit logging
 */
export interface TraceabilityService {
  // Snapshot Management
  createSnapshot(docId: DocumentId): Promise<SnapshotId>;
  getSnapshot(id: SnapshotId): DocumentSnapshot | null;
  listSnapshots(docId: DocumentId): DocumentSnapshot[];
  deleteSnapshot(id: SnapshotId): void;
  compareSnapshots(id1: SnapshotId, id2: SnapshotId): Diff | null;

  // Change History
  getReferenceHistory(refId: ReferenceId): DataReference['history'];
  getCellHistory(cellId: CellId): CellChange[];

  // Dependency Analysis
  getDependents(cellId: CellId): ReferenceId[];
  getDependencies(docId: DocumentId): DataSource[];

  // Audit Logging
  logAction(action: AuditAction): void;
  getAuditLog(filters?: AuditFilters): AuditLog[];
  clearAuditLog(): void;
}

/**
 * Default Implementation of TraceabilityService
 */
export class DefaultTraceabilityService implements TraceabilityService {
  private snapshots = new Map<SnapshotId, DocumentSnapshot>();
  private referenceHistory = new Map<ReferenceId, DataReference['history']>();
  private cellHistory = new Map<CellId, CellChange[]>();
  private auditLogs: AuditLog[] = [];

  constructor(
    private getDocumentContent: (docId: DocumentId) => any,
    private getDocumentReferences: (docId: DocumentId) => DataReference[]
  ) {
    this.setupEventListeners();
  }

  // ==================== Snapshot Management ====================

  /**
   * Create a snapshot of a document
   */
  async createSnapshot(docId: DocumentId): Promise<SnapshotId> {
    const snapshotId = `snapshot-${crypto.randomUUID()}`;

    const content = this.getDocumentContent(docId);
    const references = this.getDocumentReferences(docId);
    const checksum = this.calculateChecksum(content, references);

    const snapshot: DocumentSnapshot = {
      id: snapshotId,
      documentId: docId,
      timestamp: Date.now(),
      data: {
        content: JSON.parse(JSON.stringify(content)),
        references: JSON.parse(JSON.stringify(references)),
      },
      checksum,
    };

    this.snapshots.set(snapshotId, snapshot);

    // Log action
    this.logAction({
      action: 'snapshot_created',
      userId: 'system',
      metadata: { snapshotId, documentId: docId },
    });

    // Emit event
    eventBus.emit(Events.ReferenceCreated, { snapshotId, documentId: docId });

    return snapshotId;
  }

  /**
   * Get a specific snapshot
   */
  getSnapshot(id: SnapshotId): DocumentSnapshot | null {
    return this.snapshots.get(id) || null;
  }

  /**
   * List all snapshots for a document
   */
  listSnapshots(docId: DocumentId): DocumentSnapshot[] {
    const snapshots: DocumentSnapshot[] = [];

    this.snapshots.forEach((snapshot) => {
      if (snapshot.documentId === docId) {
        snapshots.push(snapshot);
      }
    });

    return snapshots.sort((a, b) => a.timestamp - b.timestamp);
  }

  /**
   * Delete a snapshot
   */
  deleteSnapshot(id: SnapshotId): void {
    this.snapshots.delete(id);
    this.logAction({
      action: 'snapshot_restored',
      userId: 'system',
      metadata: { snapshotId: id },
    });
  }

  /**
   * Compare two snapshots and return differences
   */
  compareSnapshots(id1: SnapshotId, id2: SnapshotId): Diff | null {
    const snapshot1 = this.snapshots.get(id1);
    const snapshot2 = this.snapshots.get(id2);

    if (!snapshot1 || !snapshot2) {
      return null;
    }

    const refs1 = snapshot1.data.references || [];
    const refs2 = snapshot2.data.references || [];

    const diff: Diff = {
      added: [],
      removed: [],
      modified: [],
    };

    // Compare references
    const refIds1 = new Set(refs1.map((r) => r.id));
    const refIds2 = new Set(refs2.map((r) => r.id));

    refIds1.forEach((refId) => {
      if (!refIds2.has(refId)) {
        diff.removed.push(refId);
      }
    });

    refIds2.forEach((refId) => {
      if (!refIds1.has(refId)) {
        diff.added.push(refId);
      }
    });

    // Find modified references
    const commonRefIds = Array.from(refIds1).filter((id) => refIds2.has(id));
    commonRefIds.forEach((refId) => {
      const ref1 = refs1.find((r) => r.id === refId);
      const ref2 = refs2.find((r) => r.id === refId);

      if (ref1 && ref2) {
        const changed = this.hasChanged(ref1, ref2);
        if (changed) {
          diff.modified.push({
            path: refId,
            oldValue: ref1.display.value,
            newValue: ref2.display.value,
          });
        }
      }
    });

    return diff;
  }

  // ==================== Change History ====================

  /**
   * Get history of a reference
   */
  getReferenceHistory(refId: ReferenceId): DataReference['history'] {
    return this.referenceHistory.get(refId) || [];
  }

  /**
   * Get history of a cell
   */
  getCellHistory(cellId: CellId): CellChange[] {
    return this.cellHistory.get(cellId) || [];
  }

  /**
   * Update reference history
   */
  updateReferenceHistory(refId: ReferenceId, history: DataReference['history']): void {
    this.referenceHistory.set(refId, history);
  }

  /**
   * Record a cell change
   */
  recordCellChange(cellId: CellId, oldValue: any, newValue: any, userId: string): void {
    const history = this.cellHistory.get(cellId) || [];
    history.push({
      timestamp: Date.now(),
      oldValue,
      newValue,
      userId,
    });
    this.cellHistory.set(cellId, history);
  }

  // ==================== Dependency Analysis ====================

  /**
   * Get all references that depend on a specific cell
   */
  getDependents(cellId: CellId): ReferenceId[] {
    const dependents: ReferenceId[] = [];

    this.referenceHistory.forEach((history, refId) => {
      const latestChange = history[history.length - 1];
      if (latestChange && latestChange.value?.toString().includes(cellId)) {
        dependents.push(refId);
      }
    });

    return dependents;
  }

  /**
   * Get all data sources referenced by a document
   */
  getDependencies(docId: DocumentId): DataSource[] {
    const snapshots = this.listSnapshots(docId);
    if (snapshots.length === 0) {
      return [];
    }

    const latestSnapshot = snapshots[snapshots.length - 1];
    const references = latestSnapshot.data.references || [];
    return references.map((ref) => ref.source);
  }

  // ==================== Audit Logging ====================

  /**
   * Log an audit action
   */
  logAction(action: AuditAction): void {
    const log: AuditLog = {
      id: crypto.randomUUID(),
      timestamp: Date.now(),
      action: action.action,
      userId: action.userId || 'system',
      metadata: action.metadata || {},
    };

    this.auditLogs.push(log);

    // Keep only last 1000 logs to prevent memory issues
    if (this.auditLogs.length > 1000) {
      this.auditLogs = this.auditLogs.slice(-1000);
    }
  }

  /**
   * Get audit logs with optional filters
   */
  getAuditLog(filters: AuditFilters = {}): AuditLog[] {
    let logs = [...this.auditLogs];

    if (filters.startDate) {
      logs = logs.filter((log) => log.timestamp >= filters.startDate!);
    }

    if (filters.endDate) {
      logs = logs.filter((log) => log.timestamp <= filters.endDate!);
    }

    if (filters.action) {
      logs = logs.filter((log) => log.action === filters.action);
    }

    if (filters.userId) {
      logs = logs.filter((log) => log.userId === filters.userId);
    }

    if (filters.limit) {
      logs = logs.slice(-filters.limit!);
    }

    return logs.sort((a, b) => b.timestamp - a.timestamp);
  }

  /**
   * Clear all audit logs
   */
  clearAuditLog(): void {
    this.auditLogs = [];
  }

  // ==================== Private Helpers ====================

  /**
   * Calculate a checksum for snapshot data
   */
  private calculateChecksum(content: any, references: DataReference[]): string {
    const data = JSON.stringify({ content, references });
    let hash = 0;
    for (let i = 0; i < data.length; i++) {
      const char = data.charCodeAt(i);
      hash = (hash << 5) - hash + char;
      hash = hash & hash; // Convert to 32-bit integer
    }
    return hash.toString(16);
  }

  /**
   * Check if two references have different values
   */
  private hasChanged(ref1: DataReference, ref2: DataReference): boolean {
    return JSON.stringify(ref1.display.value) !== JSON.stringify(ref2.display.value);
  }

  /**
   * Setup event listeners for automatic logging
   */
  private setupEventListeners(): void {
    eventBus.on(Events.ReferenceCreated, (data: any) => {
      const reference = data.reference;
      if (reference) {
        this.logAction({
          action: 'reference_created',
          userId: 'system',
          metadata: { referenceId: reference.id, documentId: reference.target?.documentId },
        });
      }
    });

    eventBus.on(Events.ReferenceUpdated, (data: any) => {
      const referenceId = data.referenceId;
      if (referenceId) {
        this.logAction({
          action: 'reference_updated',
          userId: 'system',
          metadata: { referenceId },
        });
      }
    });

    eventBus.on(Events.ReferenceBroken, (data: any) => {
      const referenceId = data.referenceId;
      if (referenceId) {
        this.logAction({
          action: 'reference_updated',
          userId: 'system',
          metadata: { referenceId, state: 'broken' },
        });
      }
    });

    eventBus.on(Events.DataChanged, (data: any) => {
      this.logAction({
        action: 'data_changed',
        userId: 'system',
        metadata: { changeCount: Array.isArray(data) ? data.length : 0 },
      });
    });

    eventBus.on(Events.DocumentSaved, (data: any) => {
      const documentId = data.documentId;
      if (documentId) {
        this.logAction({
          action: 'document_saved',
          userId: 'system',
          metadata: { documentId },
        });
      }
    });
  }

  /**
   * Cleanup resources
   */
  destroy(): void {
    this.snapshots.clear();
    this.referenceHistory.clear();
    this.cellHistory.clear();
    this.auditLogs = [];
  }
}

/**
 * Factory function to create a TraceabilityService instance
 */
export function createTraceabilityService(
  getDocumentContent: (docId: DocumentId) => any,
  getDocumentReferences: (docId: DocumentId) => DataReference[]
): TraceabilityService {
  return new DefaultTraceabilityService(getDocumentContent, getDocumentReferences);
}
