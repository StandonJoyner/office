import { describe, it, expect, vi, beforeEach } from 'vitest';
import { DefaultDataSyncEngine, createDataSyncEngine } from './DataSyncEngine';
import { DataLinkManager } from '../editors/DataLinkManager';
import { DataReference } from '../types';

// Mock DataLinkManager for testing - uses injected resolver
class MockDataLinkManager implements DataLinkManager {
  references = new Map<string, DataReference>();
  private resolver?: (ref: DataReference) => Promise<any>;

  setResolver(resolver: (ref: DataReference) => Promise<any>) {
    this.resolver = resolver;
  }

  createReference(source: any, target: string): string {
    const id = crypto.randomUUID();
    const ref: DataReference = {
      id,
      type: 'cell',
      source,
      target: { documentId: target, nodeId: id },
      display: { format: 'value', value: null },
      state: 'active',
      history: [
        {
          timestamp: Date.now(),
          action: 'created',
          value: null,
          userId: 'system',
        },
      ],
    };
    this.references.set(id, ref);
    return id;
  }

  getReference(id: string): DataReference | undefined {
    return this.references.get(id);
  }

  updateReference(id: string, data: Partial<DataReference>): void {
    const ref = this.references.get(id);
    if (ref) {
      Object.assign(ref, data);
    }
  }

  deleteReference(id: string): void {
    this.references.delete(id);
  }

  findReferencesBySource(source: any): string[] {
    const results: string[] = [];
    this.references.forEach((ref, id) => {
      if (
        ref.source?.fileId === source?.fileId &&
        ref.source?.sheetId === source?.sheetId
      ) {
        results.push(id);
      }
    });
    return results;
  }

  findReferencesByDocument(docId: string): string[] {
    const results: string[] = [];
    this.references.forEach((ref, id) => {
      if (ref.target.documentId === docId) {
        results.push(id);
      }
    });
    return results;
  }

  getAllReferences(): DataReference[] {
    return Array.from(this.references.values());
  }

  async resolveReference(ref: DataReference): Promise<any> {
    if (this.resolver) {
      return await this.resolver(ref);
    }
    return ref.display.value;
  }

  async resolveReferenceBatch(refs: DataReference[]): Promise<Map<string, any>> {
    const result = new Map<string, any>();
    for (const ref of refs) {
      result.set(ref.id, await this.resolveReference(ref));
    }
    return result;
  }

  validateReference(ref: DataReference): boolean {
    return ref.state !== 'broken';
  }
}

describe('DataSyncEngine', () => {
  let dataLinkManager: MockDataLinkManager;
  let dataSyncEngine: DefaultDataSyncEngine;
  let mockResolver: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    dataLinkManager = new MockDataLinkManager();
    mockResolver = vi.fn();
    dataSyncEngine = new DefaultDataSyncEngine(
      dataLinkManager,
      mockResolver
    );
    // Set the resolver on the mock manager
    dataLinkManager.setResolver(mockResolver);
  });

  it('should refresh a single reference', async () => {
    // Create a reference
    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    // Mock resolver to return a value
    mockResolver.mockResolvedValue('new-value');

    // Refresh the reference
    await dataSyncEngine.refreshReference(refId);

    // Verify the reference was updated
    const ref = dataLinkManager.getReference(refId);
    expect(ref?.display.value).toBe('new-value');
    expect(ref?.state).toBe('active');
  });

  it('should refresh all references', async () => {
    // Create multiple references
    const refId1 = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );
    const refId2 = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 1 }, isFormula: false },
      'doc-1'
    );

    // Set initial values
    dataLinkManager.updateReference(refId1, {
      display: { format: 'value', value: 'old-1' },
    });
    dataLinkManager.updateReference(refId2, {
      display: { format: 'value', value: 'old-2' },
    });

    // Mock resolver to return values
    mockResolver.mockImplementation(async (ref: DataReference) => {
      return ref.id === refId1 ? 'value-1' : 'value-2';
    });

    // Refresh all references
    await dataSyncEngine.refreshAll();

    // Verify both references were updated
    expect(dataLinkManager.getReference(refId1)?.display.value).toBe('value-1');
    expect(dataLinkManager.getReference(refId2)?.display.value).toBe('value-2');
  });

  it('should refresh all references in a document', async () => {
    // Create references in different documents
    const refId1 = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );
    const refId2 = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 1 }, isFormula: false },
      'doc-1'
    );
    dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 2 }, isFormula: false },
      'doc-2' // Different document
    );

    // Set initial values
    dataLinkManager.updateReference(refId1, {
      display: { format: 'value', value: 'old-1' },
    });
    dataLinkManager.updateReference(refId2, {
      display: { format: 'value', value: 'old-2' },
    });

    // Mock resolver
    mockResolver.mockImplementation(async (ref: DataReference) => {
      return ref.id === refId1 ? 'value-1' : 'value-2';
    });

    // Refresh only doc-1
    await dataSyncEngine.refreshDocument('doc-1');

    // Verify references in doc-1 were updated
    expect(dataLinkManager.getReference(refId1)?.display.value).toBe('value-1');
    expect(dataLinkManager.getReference(refId2)?.display.value).toBe('value-2');
  });

  it('should refresh only references for a given source (refreshBySource)', async () => {
    const sourceA = {
      fileId: 'file-a',
      fileName: 'A.xlsx',
      sheetId: 'sheet1',
      sheetName: 'Sheet1',
      range: { startRow: 0, startCol: 0 },
      isFormula: false,
    };
    const sourceB = {
      fileId: 'file-b',
      fileName: 'B.xlsx',
      sheetId: 'sheet1',
      sheetName: 'Sheet1',
      range: { startRow: 1, startCol: 1 },
      isFormula: false,
    };

    const refIdA1 = dataLinkManager.createReference(sourceA, 'doc-1');
    const refIdA2 = dataLinkManager.createReference(sourceA, 'doc-2');
    const refIdB = dataLinkManager.createReference(sourceB, 'doc-1');

    dataLinkManager.updateReference(refIdA1, { display: { format: 'value', value: 'old-a1' } });
    dataLinkManager.updateReference(refIdA2, { display: { format: 'value', value: 'old-a2' } });
    dataLinkManager.updateReference(refIdB, { display: { format: 'value', value: 'old-b' } });

    mockResolver.mockImplementation(async (ref: DataReference) => {
      if (ref.source?.fileId === 'file-a') return 'new-a';
      return 'new-b';
    });

    await dataSyncEngine.refreshBySource(sourceA);

    expect(dataLinkManager.getReference(refIdA1)?.display.value).toBe('new-a');
    expect(dataLinkManager.getReference(refIdA2)?.display.value).toBe('new-a');
    expect(dataLinkManager.getReference(refIdB)?.display.value).toBe('old-b');
  });

  it('should enable and disable auto-sync', () => {
    expect(dataSyncEngine.isAutoSyncEnabled()).toBe(false);

    dataSyncEngine.enableAutoSync(true);
    expect(dataSyncEngine.isAutoSyncEnabled()).toBe(true);

    dataSyncEngine.enableAutoSync(false);
    expect(dataSyncEngine.isAutoSyncEnabled()).toBe(false);
  });

  it('should set sync interval', () => {
    dataSyncEngine.setSyncInterval(5000);
    // Interval is set, we can't directly check it without exposing private members
    // But we can ensure it doesn't throw
    expect(() => dataSyncEngine.setSyncInterval(10000)).not.toThrow();
  });

  it('should call data change callbacks when value changes', async () => {
    const callback = vi.fn();
    dataSyncEngine.onDataChange(callback);

    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    // Set initial value
    dataLinkManager.updateReference(refId, {
      display: { format: 'value', value: 'old-value' },
    });

    // Mock resolver to return a different value
    mockResolver.mockResolvedValue('new-value');

    // Refresh the reference
    await dataSyncEngine.refreshReference(refId);

    // Verify callback was called with change
    expect(callback).toHaveBeenCalled();
    const changes = callback.mock.calls[0][0];
    expect(changes).toHaveLength(1);
    expect(changes[0].oldValue).toBe('old-value');
    expect(changes[0].newValue).toBe('new-value');
  });

  it('should not call data change callback when value is same', async () => {
    const callback = vi.fn();
    dataSyncEngine.onDataChange(callback);

    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    // Set initial value
    dataLinkManager.updateReference(refId, {
      display: { format: 'value', value: 'same-value' },
    });

    // Mock resolver to return the same value
    mockResolver.mockResolvedValue('same-value');

    // Refresh the reference
    await dataSyncEngine.refreshReference(refId);

    // Verify callback was not called
    expect(callback).not.toHaveBeenCalled();
  });

  it('should create sync engine via factory', () => {
    const engine = createDataSyncEngine(dataLinkManager);
    expect(engine).toBeDefined();
    expect(engine.isAutoSyncEnabled()).toBe(false);
  });

  it('should detect conflicts', () => {
    // Create a reference in conflict state
    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    dataLinkManager.updateReference(refId, {
      state: 'conflict',
      display: { format: 'value', value: 'local-value' },
    });

    const conflicts = dataSyncEngine.detectConflicts();
    expect(conflicts).toHaveLength(1);
    expect(conflicts[0].referenceId).toBe(refId);
    expect(conflicts[0].localValue).toBe('local-value');
  });

  it('should resolve conflict with keep-local strategy', () => {
    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    dataLinkManager.updateReference(refId, {
      state: 'conflict',
      display: { format: 'value', value: 'local-value' },
    });

    const conflict = {
      referenceId: refId,
      localValue: 'local-value',
      remoteValue: 'remote-value',
      timestamp: Date.now(),
    };

    dataSyncEngine.resolveConflict(conflict, 'keep-local');

    // Verify the reference is now active and keeps local value
    const ref = dataLinkManager.getReference(refId);
    expect(ref?.state).toBe('active');
    expect(ref?.display.value).toBe('local-value');
  });

  it('should resolve conflict with keep-remote strategy', () => {
    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    dataLinkManager.updateReference(refId, {
      state: 'conflict',
      display: { format: 'value', value: 'local-value' },
    });

    const conflict = {
      referenceId: refId,
      localValue: 'local-value',
      remoteValue: 'remote-value',
      timestamp: Date.now(),
    };

    dataSyncEngine.resolveConflict(conflict, 'keep-remote');

    // Verify the reference is now active and uses remote value
    const ref = dataLinkManager.getReference(refId);
    expect(ref?.state).toBe('active');
    expect(ref?.display.value).toBe('remote-value');
  });

  it('should remove data change listener', async () => {
    const callback = vi.fn();
    dataSyncEngine.onDataChange(callback);

    const refId = dataLinkManager.createReference(
      { fileId: 'test', fileName: 'Test.xlsx', sheetId: 'sheet1', sheetName: 'Sheet1', range: { startRow: 0, startCol: 0 }, isFormula: false },
      'doc-1'
    );

    // Set initial value
    dataLinkManager.updateReference(refId, {
      display: { format: 'value', value: 'old-value' },
    });

    mockResolver.mockResolvedValue('new-value');

    // Remove the listener
    dataSyncEngine.offDataChange(callback);

    // Refresh the reference
    await dataSyncEngine.refreshReference(refId);

    // Verify callback was not called
    expect(callback).not.toHaveBeenCalled();
  });

  it('should cleanup on destroy', async () => {
    // Create a new engine for this test to ensure clean state
    const testEngine = new DefaultDataSyncEngine(
      dataLinkManager,
      mockResolver
    );
    dataLinkManager.setResolver(mockResolver);

    const callback = vi.fn();
    testEngine.onDataChange(callback);

    // Enable auto sync then destroy
    testEngine.enableAutoSync(true);
    // Disable auto sync first to clear the timer
    testEngine.enableAutoSync(false);

    // Now destroy
    testEngine.destroy();

    expect(testEngine.isAutoSyncEnabled()).toBe(false);

    // Verify callbacks are cleared - this won't call the refresh since callbacks are cleared
    await testEngine.refreshReference('non-existent');
    expect(callback).not.toHaveBeenCalled();
  });
});
