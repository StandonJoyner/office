/**
 * Batch Update Hook
 *
 * Provides utilities for batch data updates with progress tracking
 */

'use client';

import { useState, useCallback, useRef } from 'react';
import { DataLinkManager, createDataLinkManager } from '../../core/editors/DataLinkManager';
import { UpdateQueue, createUpdateQueue } from '../../core/sync';

interface BatchUpdateOptions {
  batchSize?: number;
  delay?: number;
  onProgress?: (current: number, total: number) => void;
  onComplete?: (results: any[]) => void;
  onError?: (error: Error, failedItem: any) => void;
}

interface BatchUpdateItem<T> {
  id: string;
  data: T;
  updateFn: (data: T) => Promise<any>;
}

interface BatchUpdateResult<T> {
  id: string;
  data: T;
  result?: any;
  error?: Error;
  succeeded: boolean;
}

export function useBatchUpdate<T = any>(
  dataLinkManager?: DataLinkManager
) {
  const manager = dataLinkManager || createDataLinkManager();
  const updateQueue = useRef<UpdateQueue>(createUpdateQueue());

  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [results, setResults] = useState<BatchUpdateResult<T>[]>([]);

  /**
   * Execute batch update
   */
  const batchUpdate = useCallback(
    async (items: BatchUpdateItem<T>[], options: BatchUpdateOptions = {}) => {
      const {
        batchSize = 10,
        delay = 50,
        onProgress,
        onComplete,
        onError,
      } = options;

      setIsProcessing(true);
      setResults([]);
      setProgress({ current: 0, total: items.length });

      const allResults: BatchUpdateResult<T>[] = [];

      try {
        // Process in batches
        for (let i = 0; i < items.length; i += batchSize) {
          const batch = items.slice(i, i + batchSize);

          // Process batch
          const batchResults = await Promise.all(
            batch.map(async (item) => {
              try {
                const result = await item.updateFn(item.data);
                return {
                  id: item.id,
                  data: item.data,
                  result,
                  succeeded: true,
                };
              } catch (error) {
                const err = error instanceof Error ? error : new Error(String(error));
                return {
                  id: item.id,
                  data: item.data,
                  error: err,
                  succeeded: false,
                };
              }
            })
          );

          allResults.push(...batchResults);

          // Update progress
          setProgress({ current: Math.min(i + batchSize, items.length), total: items.length });
          onProgress?.(Math.min(i + batchSize, items.length), items.length);

          // Report errors
          const failedItems = batchResults.filter(r => !r.succeeded);
          for (const failed of failedItems) {
            onError?.(failed.error!, failed.data);
          }

          // Delay between batches
          if (i + batchSize < items.length) {
            await new Promise(resolve => setTimeout(resolve, delay));
          }
        }

        setResults(allResults);
        onComplete?.(allResults.filter(r => r.succeeded).map(r => r.result!));
      } catch (error) {
        const err = error instanceof Error ? error : new Error(String(error));
        onError?.(err, null);
      } finally {
        setIsProcessing(false);
      }
    },
    [manager]
  );

  /**
   * Update multiple references
   */
  const updateReferences = useCallback(
    async (
      referenceIds: string[],
      values: Map<string, any>,
      options?: BatchUpdateOptions
    ) => {
      const items: BatchUpdateItem<{ referenceId: string; value: any }>[] = referenceIds.map(id => ({
        id: id,
        data: { referenceId: id, value: values.get(id) },
        updateFn: async ({ referenceId, value }) => {
          const ref = manager.getReference(referenceId);
          if (ref) {
            manager.updateReference(referenceId, {
              display: { ...ref.display, value },
            });
          }
          return value;
        },
      }));

      await batchUpdate(items as unknown as BatchUpdateItem<T>[], options);
    },
    [manager, batchUpdate]
  );

  /**
   * Update a single reference with retry
   */
  const updateReferenceWithRetry = useCallback(
    async (
      referenceId: string,
      value: any,
      maxRetries: number = 3
    ): Promise<boolean> => {
      for (let attempt = 0; attempt < maxRetries; attempt++) {
        try {
          const ref = manager.getReference(referenceId);
          if (ref) {
            manager.updateReference(referenceId, {
              display: { ...ref.display, value },
            });
          }
          return true;
        } catch (error) {
          if (attempt === maxRetries - 1) {
            console.error(`Failed to update reference ${referenceId} after ${maxRetries} attempts:`, error);
            return false;
          }
          // Wait before retry
          await new Promise(resolve => setTimeout(resolve, 100 * (attempt + 1)));
        }
      }
      return false;
    },
    [manager]
  );

  /**
   * Clear results
   */
  const clearResults = useCallback(() => {
    setResults([]);
    setProgress({ current: 0, total: 0 });
  }, []);

  /**
   * Cancel processing
   */
  const cancel = useCallback(() => {
    // This would need actual cancellation support in UpdateQueue
    setIsProcessing(false);
  }, []);

  return {
    isProcessing,
    progress,
    results,
    batchUpdate,
    updateReferences,
    updateReferenceWithRetry,
    clearResults,
    cancel,
    get succeeded() {
      return results.filter(r => r.succeeded).length;
    },
    get failed() {
      return results.filter(r => !r.succeeded);
    },
  };
}

/**
 * Global batch update state (shared across components)
 */
const globalBatchState = {
  queue: new Map<string, any>(),
  pendingUpdates: 0,
};

/**
 * Enqueue an update for later processing
 */
export function enqueueUpdate(id: string, data: any): void {
  globalBatchState.queue.set(id, data);
  globalBatchState.pendingUpdates++;
}

/**
 * Get pending update count
 */
export function getPendingUpdateCount(): number {
  return globalBatchState.pendingUpdates;
}

/**
 * Clear all pending updates
 */
export function clearPendingUpdates(): void {
  globalBatchState.queue.clear();
  globalBatchState.pendingUpdates = 0;
}

/**
 * Get all pending updates
 */
export function getAllPendingUpdates(): Map<string, any> {
  return new Map(globalBatchState.queue);
}
