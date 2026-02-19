/**
 * Conflict Resolver for Data Synchronization
 *
 * Handles detection and resolution of conflicts between local and remote data
 */

import type { DataReference, DataChange } from '../types';

export type ConflictType = 'concurrent-edit' | 'version-mismatch' | 'data-inconsistency' | 'dependency-conflict';

export type ResolveStrategy = 'keep-local' | 'keep-remote' | 'merge' | 'manual' | 'latest-timestamp';

export interface Conflict {
  id: string;
  type: ConflictType;
  referenceId: string;
  reference?: DataReference;
  localValue: any;
  remoteValue: any;
  localTimestamp: number;
  remoteTimestamp: number;
  suggestedStrategy?: ResolveStrategy;
  resolved: boolean;
}

export interface ConflictResolution {
  conflictId: string;
  strategy: ResolveStrategy;
  resolvedValue: any;
  timestamp: number;
  userId?: string;
}

export interface MergeResult {
  value: any;
  conflicts: Conflict[];
  requiresManualResolution: boolean;
}

/**
 * Conflict Resolver
 *
 * Detects conflicts and provides resolution strategies
 */
export class ConflictResolver {
  private readonly maxConflictAge = 5 * 60 * 1000; // 5 minutes
  private readonly pendingConflicts = new Map<string, Conflict>();
  private readonly resolutionCallbacks = new Set<(conflict: Conflict) => void>();

  /**
   * Detect conflicts between local and remote data
   */
  detectConflicts(
    localData: Map<string, { value: any; timestamp: number }>,
    remoteData: Map<string, { value: any; timestamp: number }>
  ): Conflict[] {
    const conflicts: Conflict[] = [];

    // Check each reference for conflicts
    for (const [referenceId, local] of localData.entries()) {
      const remote = remoteData.get(referenceId);

      if (!remote) continue;

      // Check for concurrent edits (both modified since last sync)
      if (this.isConcurrentEdit(local, remote)) {
        conflicts.push({
          id: `conflict-${referenceId}-${Date.now()}`,
          type: 'concurrent-edit',
          referenceId,
          localValue: local.value,
          remoteValue: remote.value,
          localTimestamp: local.timestamp,
          remoteTimestamp: remote.timestamp,
          suggestedStrategy: 'latest-timestamp',
          resolved: false,
        });
      }

      // Check for version mismatch
      else if (this.isVersionMismatch(local, remote)) {
        conflicts.push({
          id: `conflict-${referenceId}-${Date.now()}`,
          type: 'version-mismatch',
          referenceId,
          localValue: local.value,
          remoteValue: remote.value,
          localTimestamp: local.timestamp,
          remoteTimestamp: remote.timestamp,
          suggestedStrategy: 'keep-remote',
          resolved: false,
        });
      }

      // Check for data inconsistency
      else if (this.isDataInconsistency(local.value, remote.value)) {
        conflicts.push({
          id: `conflict-${referenceId}-${Date.now()}`,
          type: 'data-inconsistency',
          referenceId,
          localValue: local.value,
          remoteValue: remote.value,
          localTimestamp: local.timestamp,
          remoteTimestamp: remote.timestamp,
          suggestedStrategy: 'manual',
          resolved: false,
        });
      }
    }

    return conflicts;
  }

  /**
   * Check if two updates represent concurrent edits
   */
  private isConcurrentEdit(
    local: { value: any; timestamp: number },
    remote: { value: any; timestamp: number }
  ): boolean {
    const valuesDiffer = JSON.stringify(local.value) !== JSON.stringify(remote.value);
    const bothRecentlyModified =
      (Date.now() - local.timestamp < this.maxConflictAge) &&
      (Date.now() - remote.timestamp < this.maxConflictAge);

    return valuesDiffer && bothRecentlyModified;
  }

  /**
   * Check if there's a version mismatch
   */
  private isVersionMismatch(
    local: { value: any; timestamp: number },
    remote: { value: any; timestamp: number }
  ): boolean {
    return remote.timestamp > local.timestamp &&
      JSON.stringify(local.value) !== JSON.stringify(remote.value);
  }

  /**
   * Check if data is structurally inconsistent
   */
  private isDataInconsistency(localValue: any, remoteValue: any): boolean {
    // Check type mismatches
    if (typeof localValue !== typeof remoteValue) {
      return true;
    }

    // Check for null/undefined mismatches
    if ((localValue === null || localValue === undefined) !==
        (remoteValue === null || remoteValue === undefined)) {
      return true;
    }

    // Check for array vs object mismatches
    if (Array.isArray(localValue) !== Array.isArray(remoteValue)) {
      return true;
    }

    return false;
  }

  /**
   * Resolve a conflict with the specified strategy
   */
  resolveConflict(conflict: Conflict, strategy: ResolveStrategy): any {
    const resolved = this.applyStrategy(conflict, strategy);

    // Mark conflict as resolved
    const updated: Conflict = {
      ...conflict,
      resolved: true,
    };
    this.pendingConflicts.set(conflict.id, updated);

    return resolved;
  }

  /**
   * Apply resolution strategy to get resolved value
   */
  private applyStrategy(conflict: Conflict, strategy: ResolveStrategy): any {
    switch (strategy) {
      case 'keep-local':
        return conflict.localValue;

      case 'keep-remote':
        return conflict.remoteValue;

      case 'latest-timestamp':
        return conflict.localTimestamp >= conflict.remoteTimestamp
          ? conflict.localValue
          : conflict.remoteValue;

      case 'merge':
        return this.mergeValues(conflict.localValue, conflict.remoteValue);

      case 'manual':
        // Return a placeholder for manual resolution
        return conflict.remoteValue;

      default:
        // Default to latest timestamp
        return conflict.localTimestamp >= conflict.remoteTimestamp
          ? conflict.localValue
          : conflict.remoteValue;
    }
  }

  /**
   * Merge two values
   * Basic merge strategy that can be extended
   */
  private mergeValues(local: any, remote: any): any {
    // If values are the same, return either
    if (JSON.stringify(local) === JSON.stringify(remote)) {
      return local;
    }

    // If either value is null/undefined, return the other
    if (local == null || local === undefined) return remote;
    if (remote == null || remote === undefined) return local;

    // Merge objects recursively
    if (typeof local === 'object' && !Array.isArray(local) &&
        typeof remote === 'object' && !Array.isArray(remote)) {
      return this.mergeObjects(local, remote);
    }

    // For primitive types or arrays, default to remote
    return remote;
  }

  /**
   * Merge two objects recursively
   */
  private mergeObjects(local: any, remote: any): any {
    const merged = { ...local };

    for (const key of Object.keys(remote)) {
      if (!(key in local)) {
        merged[key] = remote[key];
      } else if (typeof local[key] === 'object' && typeof remote[key] === 'object') {
        merged[key] = this.mergeObjects(local[key], remote[key]);
      } else {
        // Keep remote value for conflicting keys
        merged[key] = remote[key];
      }
    }

    return merged;
  }

  /**
   * Batch resolve multiple conflicts with the same strategy
   */
  resolveConflicts(conflicts: Conflict[], strategy: ResolveStrategy): ConflictResolution[] {
    return conflicts.map(conflict => ({
      conflictId: conflict.id,
      strategy,
      resolvedValue: this.applyStrategy(conflict, strategy),
      timestamp: Date.now(),
    }));
  }

  /**
   * Get suggested resolution for a conflict
   */
  getSuggestedResolution(conflict: Conflict): ResolveStrategy {
    return conflict.suggestedStrategy || 'latest-timestamp';
  }

  /**
   * Check if a conflict requires manual resolution
   */
  requiresManualResolution(conflict: Conflict): boolean {
    return conflict.type === 'data-inconsistency' ||
      conflict.suggestedStrategy === 'manual';
  }

  /**
   * Get pending conflicts
   */
  getPendingConflicts(): Conflict[] {
    return Array.from(this.pendingConflicts.values()).filter(c => !c.resolved);
  }

  /**
   * Get all conflicts (including resolved)
   */
  getAllConflicts(): Conflict[] {
    return Array.from(this.pendingConflicts.values());
  }

  /**
   * Register a callback for conflict resolution
   */
  onConflictResolved(callback: (conflict: Conflict) => void): () => void {
    this.resolutionCallbacks.add(callback);

    return () => {
      this.resolutionCallbacks.delete(callback);
    };
  }

  /**
   * Clear all conflicts
   */
  clear(): void {
    this.pendingConflicts.clear();
  }

  /**
   * Clean up old resolved conflicts
   */
  cleanupOldConflicts(maxAge: number = 24 * 60 * 60 * 1000): number {
    const now = Date.now();
    let cleaned = 0;

    for (const [id, conflict] of this.pendingConflicts.entries()) {
      if (conflict.resolved && (now - conflict.remoteTimestamp) > maxAge) {
        this.pendingConflicts.delete(id);
        cleaned++;
      }
    }

    return cleaned;
  }

  /**
   * Get conflict statistics
   */
  getStats() {
    const conflicts = Array.from(this.pendingConflicts.values());

    return {
      total: conflicts.length,
      resolved: conflicts.filter(c => c.resolved).length,
      unresolved: conflicts.filter(c => !c.resolved).length,
      byType: {
        'concurrent-edit': conflicts.filter(c => c.type === 'concurrent-edit').length,
        'version-mismatch': conflicts.filter(c => c.type === 'version-mismatch').length,
        'data-inconsistency': conflicts.filter(c => c.type === 'data-inconsistency').length,
        'dependency-conflict': conflicts.filter(c => c.type === 'dependency-conflict').length,
      },
    };
  }

  /**
   * Destroy the conflict resolver
   */
  destroy(): void {
    this.pendingConflicts.clear();
    this.resolutionCallbacks.clear();
  }
}

export function createConflictResolver(): ConflictResolver {
  return new ConflictResolver();
}
