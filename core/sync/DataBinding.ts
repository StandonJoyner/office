/**
 * Bidirectional Data Binding System
 *
 * Provides reactive two-way binding between Excel data sources and Word document references
 */

import { eventBus, Events } from '../eventBus';
import type { DataReference, ReferenceId, DataSource } from '../types';

export type BindingType = 'cell' | 'range' | 'named' | 'formula';

export interface Binding {
  id: string;
  referenceId: ReferenceId;
  dataSource: DataSource;
  bindingType: BindingType;
  lastSyncAt: number;
  lastSyncValue: any;
  pendingUpdate: boolean;
  listeners: Set<(value: any) => void>;
}

export interface BindingEvent {
  bindingId: string;
  referenceId: ReferenceId;
  type: 'value-changed' | 'source-updated' | 'target-updated' | 'sync-start' | 'sync-end';
  oldValue?: any;
  newValue?: any;
  timestamp: number;
}

export interface BindingOptions {
  autoSync?: boolean;
  syncDelay?: number; // ms
  syncOnError?: boolean;
  batchSize?: number;
}

const DEFAULT_OPTIONS: BindingOptions = {
  autoSync: true,
  syncDelay: 100,
  syncOnError: true,
  batchSize: 10,
};

/**
 * Signal-like class for reactive values
 */
class Signal<T> {
  private value: T;
  private listeners = new Set<(value: T) => void>();

  constructor(initialValue: T) {
    this.value = initialValue;
  }

  get(): T {
    return this.value;
  }

  set(newValue: T): void {
    const oldValue = this.value;
    this.value = newValue;

    // Notify listeners
    this.listeners.forEach(listener => {
      try {
        listener(newValue);
      } catch (error) {
        console.error('Error in signal listener:', error);
      }
    });
  }

  subscribe(listener: (value: T) => void): () => void {
    this.listeners.add(listener);
    // Immediately call with current value
    listener(this.value);

    // Return unsubscribe function
    return () => {
      this.listeners.delete(listener);
    };
  }

  peek(): T {
    return this.value;
  }
}

/**
 * Computed signal that derives from other signals
 */
class ComputedSignal<T> {
  private value: T | null = null;
  private listeners = new Set<(value: T) => void>();
  private readonly dependencies: Set<Signal<any>>;
  private isComputing = false;

  constructor(
    dependencies: Signal<any>[],
    private readonly computeFn: () => T
  ) {
    this.dependencies = new Set(dependencies);

    // Subscribe to dependencies
    dependencies.forEach(dep => {
      dep.subscribe(() => {
        if (!this.isComputing) {
          this.recompute();
        }
      });
    });

    // Initial compute
    this.recompute();
  }

  private recompute(): void {
    if (this.isComputing) return;

    this.isComputing = true;

    try {
      const newValue = this.computeFn();

      if (this.value !== newValue) {
        this.value = newValue;

        // Notify listeners
        this.listeners.forEach(listener => {
          try {
            listener(newValue);
          } catch (error) {
            console.error('Error in computed signal listener:', error);
          }
        });
      }
    } finally {
      this.isComputing = false;
    }
  }

  get(): T {
    if (this.value === null) {
      this.recompute();
    }
    return this.value as T;
  }

  subscribe(listener: (value: T) => void): () => void {
    this.listeners.add(listener);
    listener(this.get());

    return () => {
      this.listeners.delete(listener);
    };
  }
}

/**
 * Bidirectional Data Binding Manager
 *
 * Manages bindings between data sources and references
 */
export class DataBindingManager {
  private readonly bindings = new Map<ReferenceId, Binding>();
  private readonly signals = new Map<ReferenceId, Signal<any>>();
  private readonly computedSignals = new Map<string, ComputedSignal<any>>();
  private readonly bindingEventCallbacks = new Set<(event: BindingEvent) => void>();
  private readonly syncQueues = new Map<ReferenceId, NodeJS.Timeout>();
  private readonly options: BindingOptions;

  constructor(options: Partial<BindingOptions> = {}) {
    this.options = { ...DEFAULT_OPTIONS, ...options };

    // Listen to reference updates from event bus
    this.setupEventListeners();
  }

  /**
   * Create a binding for a reference
   */
  createBinding(reference: DataReference, options?: Partial<BindingOptions>): Binding {
    const binding: Binding = {
      id: crypto.randomUUID(),
      referenceId: reference.id,
      dataSource: reference.source,
      bindingType: reference.type,
      lastSyncAt: reference.metadata?.updated ?? Date.now(),
      lastSyncValue: reference.display.value,
      pendingUpdate: false,
      listeners: new Set(),
    };

    this.bindings.set(reference.id, binding);

    // Create a signal for this binding
    const signal = new Signal(reference.display.value);
    this.signals.set(reference.id, signal);

    // Emit binding creation event
    this.emitBindingEvent({
      bindingId: binding.id,
      referenceId: reference.id,
      type: 'sync-end',
      newValue: reference.display.value,
      timestamp: Date.now(),
    });

    return binding;
  }

  /**
   * Update the value of a binding
   */
  updateBinding(referenceId: ReferenceId, value: any, source: 'local' | 'remote' = 'remote'): void {
    const binding = this.bindings.get(referenceId);
    const signal = this.signals.get(referenceId);

    if (!binding || !signal) {
      console.warn(`Binding not found for reference ${referenceId}`);
      return;
    }

    const oldValue = signal.get();

    // Only update if value changed
    if (JSON.stringify(oldValue) !== JSON.stringify(value)) {
      // Update signal
      signal.set(value);

      // Update binding metadata
      binding.lastSyncValue = value;
      binding.lastSyncAt = Date.now();
      binding.pendingUpdate = false;

      // Emit binding update event
      this.emitBindingEvent({
        bindingId: binding.id,
        referenceId,
        type: source === 'remote' ? 'source-updated' : 'target-updated',
        oldValue,
        newValue: value,
        timestamp: Date.now(),
      });

      // Notify binding listeners
      binding.listeners.forEach(listener => {
        try {
          listener(value);
        } catch (error) {
          console.error('Error in binding listener:', error);
        }
      });

      // Trigger auto-sync if enabled
      if (this.options.autoSync && source === 'remote') {
        this.scheduleSync(referenceId);
      }
    }
  }

  /**
   * Get the current value of a binding
   */
  getBindingValue(referenceId: ReferenceId): any {
    const signal = this.signals.get(referenceId);
    return signal ? signal.get() : undefined;
  }

  /**
   * Subscribe to value changes for a binding
   */
  subscribe(referenceId: ReferenceId, listener: (value: any) => void): () => void {
    const binding = this.bindings.get(referenceId);
    const signal = this.signals.get(referenceId);

    if (!binding || !signal) {
      console.warn(`Binding not found for reference ${referenceId}`);
      return () => {};
    }

    // Add to binding listeners
    binding.listeners.add(listener);

    // Subscribe to signal
    const unsubscribe = signal.subscribe(listener);

    // Return combined unsubscribe
    return () => {
      binding.listeners.delete(listener);
      unsubscribe();
    };
  }

  /**
   * Create a computed binding that derives from other bindings
   */
  createComputedBinding(
    id: string,
    dependencyIds: ReferenceId[],
    computeFn: () => any
  ): ComputedSignal<any> {
    const dependencySignals = dependencyIds
      .map(refId => this.signals.get(refId))
      .filter((s): s is Signal<any> => s !== undefined);

    const computed = new ComputedSignal(dependencySignals, computeFn);
    this.computedSignals.set(id, computed);

    return computed;
  }

  /**
   * Remove a binding
   */
  removeBinding(referenceId: ReferenceId): void {
    const binding = this.bindings.get(referenceId);

    if (binding) {
      // Clear sync timer
      const timer = this.syncQueues.get(referenceId);
      if (timer) {
        clearTimeout(timer);
        this.syncQueues.delete(referenceId);
      }

      // Clean up signal
      this.signals.delete(referenceId);

      // Remove binding
      this.bindings.delete(referenceId);

      // Emit binding removal event
      this.emitBindingEvent({
        bindingId: binding.id,
        referenceId,
        type: 'sync-end',
        timestamp: Date.now(),
      });
    }
  }

  /**
   * Schedule a sync for a binding (debounced)
   */
  private scheduleSync(referenceId: ReferenceId): void {
    // Clear existing timer
    const existingTimer = this.syncQueues.get(referenceId);
    if (existingTimer) {
      clearTimeout(existingTimer);
    }

    // Schedule new sync
    const timer = setTimeout(() => {
      this.performSync(referenceId);
      this.syncQueues.delete(referenceId);
    }, this.options.syncDelay || 100);

    this.syncQueues.set(referenceId, timer);
  }

  /**
   * Perform sync for a binding
   */
  private async performSync(referenceId: ReferenceId): Promise<void> {
    const binding = this.bindings.get(referenceId);

    if (!binding) return;

    binding.pendingUpdate = true;

    // Emit sync start event
    this.emitBindingEvent({
      bindingId: binding.id,
      referenceId,
      type: 'sync-start',
      timestamp: Date.now(),
    });

    // Trigger reference update through event bus
    eventBus.emit(Events.DataChanged, [{
      referenceId,
      oldValue: binding.lastSyncValue,
      newValue: this.getBindingValue(referenceId),
      timestamp: Date.now(),
      source: 'auto',
    }]);
  }

  /**
   * Batch update multiple bindings
   */
  batchUpdate(updates: Array<{ referenceId: ReferenceId; value: any; source?: 'local' | 'remote' }>): void {
    for (const update of updates) {
      this.updateBinding(update.referenceId, update.value, update.source);
    }
  }

  /**
   * Get all bindings for a data source
   */
  getBindingsByDataSource(dataSource: DataSource): Binding[] {
    return Array.from(this.bindings.values()).filter(
      binding =>
        binding.dataSource.fileId === dataSource.fileId &&
        binding.dataSource.sheetId === dataSource.sheetId
    );
  }

  /**
   * Get a binding by reference ID
   */
  getBinding(referenceId: ReferenceId): Binding | undefined {
    return this.bindings.get(referenceId);
  }

  /**
   * Get all bindings
   */
  getAllBindings(): Binding[] {
    return Array.from(this.bindings.values());
  }

  /**
   * Get pending updates
   */
  getPendingUpdates(): Binding[] {
    return Array.from(this.bindings.values()).filter(b => b.pendingUpdate);
  }

  /**
   * Mark a binding as stale
   */
  markStale(referenceId: ReferenceId): void {
    const binding = this.bindings.get(referenceId);
    if (binding) {
      binding.pendingUpdate = true;
    }
  }

  /**
   * Register a callback for binding events
   */
  onBindingEvent(callback: (event: BindingEvent) => void): () => void {
    this.bindingEventCallbacks.add(callback);

    return () => {
      this.bindingEventCallbacks.delete(callback);
    };
  }

  /**
   * Emit a binding event to all listeners
   */
  private emitBindingEvent(event: BindingEvent): void {
    this.bindingEventCallbacks.forEach(callback => {
      try {
        callback(event);
      } catch (error) {
        console.error('Error in binding event callback:', error);
      }
    });
  }

  /**
   * Set up event listeners for external changes
   */
  private setupEventListeners(): void {
    // Listen for reference updates
    eventBus.on(Events.ReferenceUpdated, (data: any) => {
      const { referenceId, newState } = data;

      if (newState?.display?.value !== undefined) {
        this.updateBinding(referenceId, newState.display.value, 'local');
      }
    });

    // Listen for data source changes
    eventBus.on(Events.DataChanged, (changes: any[]) => {
      for (const change of changes) {
        const binding = this.bindings.get(change.referenceId);

        if (binding && change.source === 'auto') {
          // Auto-sync updates from data source
          this.updateBinding(change.referenceId, change.newValue, 'remote');
        }
      }
    });
  }

  /**
   * Get binding statistics
   */
  getStats() {
    const bindings = Array.from(this.bindings.values());

    return {
      total: bindings.length,
      byType: {
        cell: bindings.filter(b => b.bindingType === 'cell').length,
        range: bindings.filter(b => b.bindingType === 'range').length,
        named: bindings.filter(b => b.bindingType === 'named').length,
        formula: bindings.filter(b => b.bindingType === 'formula').length,
      },
      pending: bindings.filter(b => b.pendingUpdate).length,
      averageSyncAge: this.getAverageSyncAge(bindings),
    };
  }

  /**
   * Calculate average sync age
   */
  private getAverageSyncAge(bindings: Binding[]): number {
    if (bindings.length === 0) return 0;

    const now = Date.now();
    const totalAge = bindings.reduce((sum, b) => {
      return sum + (now - b.lastSyncAt);
    }, 0);

    return Math.round(totalAge / bindings.length);
  }

  /**
   * Clear all bindings
   */
  clear(): void {
    // Clear all timers
    for (const timer of this.syncQueues.values()) {
      clearTimeout(timer);
    }
    this.syncQueues.clear();

    // Clear all bindings
    this.bindings.clear();
    this.signals.clear();
    this.computedSignals.clear();
  }

  /**
   * Destroy the binding manager
   */
  destroy(): void {
    this.clear();
    this.bindingEventCallbacks.clear();
  }
}

export function createDataBindingManager(options?: Partial<BindingOptions>): DataBindingManager {
  return new DataBindingManager(options);
}
