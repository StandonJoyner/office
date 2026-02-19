import { describe, it, expect, beforeEach } from 'vitest';
import { DefaultTraceabilityService, createTraceabilityService } from './TraceabilityService';
import { DataReference, DocumentSnapshot } from '../types';

// Mock document store
const mockDocumentStore = new Map<string, any>();
const mockReferenceStore = new Map<string, DataReference[]>();

describe('TraceabilityService', () => {
  let service: DefaultTraceabilityService;

  beforeEach(() => {
    mockDocumentStore.clear();
    mockReferenceStore.clear();
    service = createTraceabilityService(
      (docId) => mockDocumentStore.get(docId),
      (docId) => mockReferenceStore.get(docId) || []
    );
  });

  it('should create a snapshot', async () => {
    const docId = 'doc-1';
    mockDocumentStore.set(docId, { type: 'doc', content: 'test' });
    mockReferenceStore.set(docId, [
      {
        id: 'ref-1',
        type: 'cell',
        source: {
          fileId: 'file-1',
          fileName: 'Test.xlsx',
          sheetId: 'sheet1',
          sheetName: 'Sheet1',
          range: { startRow: 0, startCol: 0 },
          isFormula: false,
        },
        target: { documentId: docId, nodeId: 'node-1' },
        display: { format: 'value', value: '100' },
        state: 'active',
        history: [],
      },
    ]);

    const snapshotId = await service.createSnapshot(docId);

    expect(snapshotId).toBeDefined();
    expect(snapshotId).toContain('snapshot-');

    const snapshot = service.getSnapshot(snapshotId);
    expect(snapshot).not.toBeNull();
    expect(snapshot?.documentId).toBe(docId);
    expect(snapshot?.data.references).toHaveLength(1);
  });

  it('should list snapshots for a document', async () => {
    const docId = 'doc-1';
    mockDocumentStore.set(docId, { type: 'doc', content: 'test' });
    mockReferenceStore.set(docId, []);

    await service.createSnapshot(docId);
    await service.createSnapshot(docId);
    await service.createSnapshot(docId);

    const snapshots = service.listSnapshots(docId);
    expect(snapshots).toHaveLength(3);
  });

  it('should get null for non-existent snapshot', () => {
    const snapshot = service.getSnapshot('non-existent');
    expect(snapshot).toBeNull();
  });

  it('should delete a snapshot', async () => {
    const docId = 'doc-1';
    mockDocumentStore.set(docId, { type: 'doc', content: 'test' });
    mockReferenceStore.set(docId, []);

    const snapshotId = await service.createSnapshot(docId);
    expect(service.getSnapshot(snapshotId)).not.toBeNull();

    service.deleteSnapshot(snapshotId);
    expect(service.getSnapshot(snapshotId)).toBeNull();
  });

  it('should compare two snapshots', async () => {
    const docId = 'doc-1';
    mockDocumentStore.set(docId, { type: 'doc', content: 'test' });

    const ref1: DataReference = {
      id: 'ref-1',
      type: 'cell',
      source: {
        fileId: 'file-1',
        fileName: 'Test.xlsx',
        sheetId: 'sheet1',
        sheetName: 'Sheet1',
        range: { startRow: 0, startCol: 0 },
        isFormula: false,
      },
      target: { documentId: docId, nodeId: 'node-1' },
      display: { format: 'value', value: '100' },
      state: 'active',
      history: [],
    };

    const ref2: DataReference = {
      ...ref1,
      display: { format: 'value', value: '200' },
    };

    // Create first snapshot
    mockReferenceStore.set(docId, [ref1]);
    const snapshotId1 = await service.createSnapshot(docId);

    // Create second snapshot with changed value
    mockReferenceStore.set(docId, [ref2]);
    const snapshotId2 = await service.createSnapshot(docId);

    const diff = service.compareSnapshots(snapshotId1, snapshotId2);
    expect(diff).not.toBeNull();
    expect(diff?.modified).toHaveLength(1);
    expect(diff?.modified[0].oldValue).toBe('100');
    expect(diff?.modified[0].newValue).toBe('200');
  });

  it('should return null when comparing non-existent snapshots', async () => {
    const diff = service.compareSnapshots('non-existent-1', 'non-existent-2');
    expect(diff).toBeNull();
  });

  it('should get reference history', () => {
    const refId = 'ref-1';
    const history = [
      { timestamp: 100, action: 'created', value: null, userId: 'user1' },
      { timestamp: 200, action: 'updated', value: 'new', userId: 'user2' },
    ];

    service.updateReferenceHistory(refId, history);

    const retrievedHistory = service.getReferenceHistory(refId);
    expect(retrievedHistory).toEqual(history);
  });

  it('should return empty history for non-existent reference', () => {
    const history = service.getReferenceHistory('non-existent');
    expect(history).toEqual([]);
  });

  it('should record cell change', () => {
    const cellId = 'cell-1';
    service.recordCellChange(cellId, 'old', 'new', 'user1');
    service.recordCellChange(cellId, 'new', 'newer', 'user2');

    const history = service.getCellHistory(cellId);
    expect(history).toHaveLength(2);
    expect(history[0].oldValue).toBe('old');
    expect(history[0].newValue).toBe('new');
    expect(history[1].newValue).toBe('newer');
  });

  it('should return empty history for non-existent cell', () => {
    const history = service.getCellHistory('non-existent');
    expect(history).toEqual([]);
  });

  it('should create service via factory', () => {
    const factoryService = createTraceabilityService(
      (docId) => mockDocumentStore.get(docId),
      (docId) => mockReferenceStore.get(docId) || []
    );
    expect(factoryService).toBeDefined();
    expect(factoryService.listSnapshots('doc-1')).toEqual([]);
  });

  it('should log audit actions', () => {
    const initialLogCount = service.getAuditLog().length;

    service.logAction({
      action: 'document_saved',
      userId: 'user1',
      metadata: { documentId: 'doc-1' },
    });

    const logs = service.getAuditLog();
    expect(logs).toHaveLength(initialLogCount + 1);
    expect(logs[logs.length - 1].action).toBe('document_saved');
    expect(logs[logs.length - 1].userId).toBe('user1');
  });

  it('should filter audit logs', () => {
    service.logAction({
      action: 'document_saved',
      userId: 'user1',
    });
    service.logAction({
      action: 'reference_created',
      userId: 'user2',
    });
    service.logAction({
      action: 'document_saved',
      userId: 'user1',
    });

    const savedLogs = service.getAuditLog({ action: 'document_saved' });
    expect(savedLogs).toHaveLength(2);

    const user1Logs = service.getAuditLog({ userId: 'user1' });
    expect(user1Logs).toHaveLength(2);

    const user2Logs = service.getAuditLog({ userId: 'user2' });
    expect(user2Logs).toHaveLength(1);
  });

  it('should respect audit log limit', () => {
    // Clear and add many logs
    service.clearAuditLog();

    for (let i = 0; i < 15; i++) {
      service.logAction({
        action: 'document_saved',
        userId: 'user1',
      });
    }

    const allLogs = service.getAuditLog();
    expect(allLogs.length).toBeGreaterThan(10); // Should have some limit

    const limitedLogs = service.getAuditLog({ limit: 5 });
    expect(limitedLogs).toHaveLength(5);
  });

  it('should clear audit logs', () => {
    service.logAction({
      action: 'document_saved',
      userId: 'user1',
    });

    expect(service.getAuditLog()).not.toHaveLength(0);

    service.clearAuditLog();

    expect(service.getAuditLog()).toHaveLength(0);
  });

  it('should filter audit logs by date range', () => {
    const now = Date.now();
    const oneHourAgo = now - 3600000;

    service.clearAuditLog();
    service.logAction({
      action: 'document_saved',
      metadata: { timestamp: oneHourAgo },
    });

    // Manually set timestamp for test
    const logs = service.getAuditLog();
    if (logs.length > 0) {
      (logs[0] as any).timestamp = oneHourAgo;
    }

    const recentLogs = service.getAuditLog({
      startDate: oneHourAgo - 1000,
      endDate: now + 1000,
    });

    expect(recentLogs.length).toBeGreaterThan(0);
  });

  it('should cleanup on destroy', () => {
    const docId = 'doc-1';
    mockDocumentStore.set(docId, { type: 'doc', content: 'test' });
    mockReferenceStore.set(docId, []);

    service.createSnapshot(docId);
    expect(service.listSnapshots(docId)).toHaveLength(1);

    service.destroy();
    expect(service.listSnapshots(docId)).toHaveLength(0);
  });
});
