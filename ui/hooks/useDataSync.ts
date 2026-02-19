'use client';

import { useEffect, useRef, useCallback } from 'react';
import { DataSyncEngine, createDataSyncEngine } from '@/core/sync/DataSyncEngine';
import { DataLinkManager, MemoryDataLinkManager } from '@/core/editors/DataLinkManager';
import { eventBus, Events } from '@/core/eventBus';
import { DataChange } from '@/core/types';

/**
 * Data Sync Hook
 *
 * Provides access to the DataSyncEngine and handles data updates.
 */
export function useDataSync(documentId: string) {
  const dataLinkManagerRef = useRef<DataLinkManager | null>(null);
  const syncEngineRef = useRef<DataSyncEngine | null>(null);
  const dataChangeListenersRef = useRef<Set<(changes: DataChange[]) => void>>(new Set());

  // Initialize services on mount
  useEffect(() => {
    dataLinkManagerRef.current = new MemoryDataLinkManager();

    // Create sync engine with the data link manager's resolver
    // The resolver delegates to ExcelDataSourceManager which resolves from actual Excel data
    syncEngineRef.current = createDataSyncEngine(
      dataLinkManagerRef.current,
      async (ref) => dataLinkManagerRef.current!.resolveReference(ref)
    );

    return () => {
      if (syncEngineRef.current) {
        syncEngineRef.current.destroy();
      }
    };
  }, []);

  const syncEngine = syncEngineRef.current;

  /**
   * Refresh all references in the document
   */
  const refreshAll = useCallback(async () => {
    if (!syncEngine || !documentId) return;
    await syncEngine.refreshDocument(documentId);
  }, [syncEngine, documentId]);

  /**
   * Add a data change listener
   */
  const onDataChange = useCallback((callback: (changes: DataChange[]) => void) => {
    if (!syncEngine) return;

    dataChangeListenersRef.current.add(callback);
    syncEngine.onDataChange(callback);

    return () => {
      dataChangeListenersRef.current.delete(callback);
      syncEngine.offDataChange(callback);
    };
  }, [syncEngine]);

  /**
   * Refresh a single reference by ID
   */
  const refreshReference = useCallback(async (referenceId: string) => {
    if (!syncEngine) return;
    await syncEngine.refreshReference(referenceId);
  }, [syncEngine]);

  /**
   * Get the data link manager
   */
  const getDataLinkManager = useCallback(() => {
    return dataLinkManagerRef.current;
  }, []);

  return {
    syncEngine,
    dataLinkManager: dataLinkManagerRef.current,
    refreshAll,
    refreshReference,
    getDataLinkManager,
    onDataChange,
  };
}
