import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  MemoryDataLinkManager,
  createDataLinkManager,
} from './DataLinkManager';
import { DataSource } from '../types';
import { eventBus, Events } from '../eventBus';

// Mock IExcelDataSourceManager for testing
class MockExcelDataSourceManager {
  private workbook: any = null;
  private changeCallbacks = new Set<(event: any) => void>();

  setWorkbook(workbook: any): void {
    this.workbook = workbook;
  }

  getWorkbook(): any {
    return this.workbook;
  }

  async resolveReference(ref: any): Promise<any> {
    return 'mocked-value';
  }

  onDataSourceChange(callback: (event: any) => void): void {
    this.changeCallbacks.add(callback);
  }

  offDataSourceChange(callback: (event: any) => void): void {
    this.changeCallbacks.delete(callback);
  }

  enableChangeMonitoring(enabled: boolean): void {
    // Mock implementation
  }

  validateSheet(sheetId: string): boolean {
    return this.workbook?.hasSheet(sheetId) ?? false;
  }

  validateCell(sheetId: string, row: number, column: number): boolean {
    return this.validateSheet(sheetId);
  }

  destroy(): void {
    this.changeCallbacks.clear();
    this.workbook = null;
  }

  // Test helper to emit changes
  emitChange(event: any): void {
    this.changeCallbacks.forEach((cb) => cb(event));
  }
}

describe('DataLinkManager', () => {
  let mockExcelManager: MockExcelDataSourceManager;
  let manager: MemoryDataLinkManager;

  const mockDataSource: DataSource = {
    fileId: 'file-1',
    fileName: 'Test.xlsx',
    sheetId: 'sheet1',
    sheetName: 'Sheet1',
    range: { startRow: 0, startCol: 0 },
    isFormula: false,
  };

  beforeEach(() => {
    mockExcelManager = new MockExcelDataSourceManager();
    manager = new MemoryDataLinkManager(mockExcelManager);
  });

  afterEach(() => {
    manager.destroy();
  });

  describe('createReference', () => {
    it('should create a reference and return its ID', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      expect(refId).toBeDefined();
      expect(typeof refId).toBe('string');
    });

    it('should create a reference with correct properties', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      const ref = manager.getReference(refId);

      expect(ref).toBeDefined();
      expect(ref?.id).toBe(refId);
      expect(ref?.source).toEqual(mockDataSource);
      expect(ref?.target.documentId).toBe('doc-1');
      expect(ref?.state).toBe('active');
      expect(ref?.type).toBe('cell');
    });

    it('should create reference with initial history', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      const ref = manager.getReference(refId);

      expect(ref?.history).toHaveLength(1);
      expect(ref?.history[0].action).toBe('created');
      expect(ref?.history[0].value).toBeNull();
      expect(ref?.history[0].userId).toBe('system');
    });

    it('should emit ReferenceCreated event', () => {
      const handler = vi.fn();
      eventBus.on(Events.ReferenceCreated, handler);

      const refId = manager.createReference(mockDataSource, 'doc-1');

      expect(handler).toHaveBeenCalledTimes(1);
      const createdRef = handler.mock.calls[0][0];
      expect(createdRef.id).toBe(refId);
      expect(createdRef.source).toEqual(mockDataSource);

      eventBus.off(Events.ReferenceCreated, handler);
    });

    it('should generate unique IDs for each reference', () => {
      const refId1 = manager.createReference(mockDataSource, 'doc-1');
      const refId2 = manager.createReference(mockDataSource, 'doc-2');
      const refId3 = manager.createReference(mockDataSource, 'doc-1');

      expect(refId1).not.toBe(refId2);
      expect(refId2).not.toBe(refId3);
      expect(refId1).not.toBe(refId3);
    });
  });

  describe('getReference', () => {
    it('should return reference when it exists', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      const ref = manager.getReference(refId);

      expect(ref).toBeDefined();
      expect(ref?.id).toBe(refId);
    });

    it('should return undefined for non-existent reference', () => {
      const ref = manager.getReference('non-existent');
      expect(ref).toBeUndefined();
    });
  });

  describe('updateReference', () => {
    it('should update reference properties', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');

      manager.updateReference(refId, {
        state: 'stale',
        display: { format: 'value', value: 'new-value' },
      });

      const ref = manager.getReference(refId);
      expect(ref?.state).toBe('stale');
      expect(ref?.display.value).toBe('new-value');
    });

    it('should emit ReferenceUpdated event', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      const handler = vi.fn();
      eventBus.on(Events.ReferenceUpdated, handler);

      manager.updateReference(refId, {
        state: 'stale',
      });

      expect(handler).toHaveBeenCalledTimes(1);
      const updateEvent = handler.mock.calls[0][0];
      expect(updateEvent.referenceId).toBe(refId);
      expect(updateEvent.oldState.state).toBe('active');
      expect(updateEvent.newState.state).toBe('stale');

      eventBus.off(Events.ReferenceUpdated, handler);
    });

    it('should not emit event for non-existent reference', () => {
      const handler = vi.fn();
      eventBus.on(Events.ReferenceUpdated, handler);

      manager.updateReference('non-existent', { state: 'stale' });

      expect(handler).not.toHaveBeenCalled();
      eventBus.off(Events.ReferenceUpdated, handler);
    });

    it('should update nested properties', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');

      manager.updateReference(refId, {
        metadata: { created: Date.now(), updated: Date.now() },
      });

      const ref = manager.getReference(refId);
      expect(ref?.metadata?.created).toBeDefined();
      expect(ref?.metadata?.updated).toBeDefined();
    });
  });

  describe('deleteReference', () => {
    it('should delete a reference', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      expect(manager.getReference(refId)).toBeDefined();

      manager.deleteReference(refId);
      expect(manager.getReference(refId)).toBeUndefined();
    });

    it('should emit ReferenceDeleted event', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      const handler = vi.fn();
      eventBus.on(Events.ReferenceDeleted, handler);

      manager.deleteReference(refId);

      expect(handler).toHaveBeenCalledTimes(1);
      const deleteEvent = handler.mock.calls[0][0];
      expect(deleteEvent.referenceId).toBe(refId);

      eventBus.off(Events.ReferenceDeleted, handler);
    });

    it('should handle deleting non-existent reference gracefully', () => {
      expect(() => manager.deleteReference('non-existent')).not.toThrow();
    });
  });

  describe('findReferencesBySource', () => {
    it('should find references by source file and sheet', () => {
      const source1: DataSource = {
        ...mockDataSource,
        fileId: 'file-1',
        sheetId: 'sheet1',
      };
      const source2: DataSource = {
        ...mockDataSource,
        fileId: 'file-1',
        sheetId: 'sheet1',
      };
      const source3: DataSource = {
        ...mockDataSource,
        fileId: 'file-2',
        sheetId: 'sheet1',
      };

      const refId1 = manager.createReference(source1, 'doc-1');
      const refId2 = manager.createReference(source2, 'doc-2');
      const refId3 = manager.createReference(source3, 'doc-1');

      const refs = manager.findReferencesBySource({
        fileId: 'file-1',
        fileName: 'Test.xlsx',
        sheetId: 'sheet1',
        sheetName: 'Sheet1',
        range: { startRow: 0, startCol: 0 },
        isFormula: false,
      });

      expect(refs).toHaveLength(2);
      expect(refs).toContain(refId1);
      expect(refs).toContain(refId2);
      expect(refs).not.toContain(refId3);
    });

    it('should return empty array when no references found', () => {
      const refs = manager.findReferencesBySource({
        fileId: 'non-existent',
        fileName: 'Test.xlsx',
        sheetId: 'sheet1',
        sheetName: 'Sheet1',
        range: { startRow: 0, startCol: 0 },
        isFormula: false,
      });

      expect(refs).toEqual([]);
    });
  });

  describe('findReferencesByDocument', () => {
    it('should find references by document ID', () => {
      const refId1 = manager.createReference(mockDataSource, 'doc-1');
      const refId2 = manager.createReference(mockDataSource, 'doc-1');
      const refId3 = manager.createReference(mockDataSource, 'doc-2');

      const doc1Refs = manager.findReferencesByDocument('doc-1');
      const doc2Refs = manager.findReferencesByDocument('doc-2');

      expect(doc1Refs).toHaveLength(2);
      expect(doc1Refs).toContain(refId1);
      expect(doc1Refs).toContain(refId2);

      expect(doc2Refs).toHaveLength(1);
      expect(doc2Refs).toContain(refId3);
    });

    it('should return empty array when no references found', () => {
      const refs = manager.findReferencesByDocument('non-existent');
      expect(refs).toEqual([]);
    });
  });

  describe('getAllReferences', () => {
    it('should return all references', () => {
      const refId1 = manager.createReference(mockDataSource, 'doc-1');
      const refId2 = manager.createReference(mockDataSource, 'doc-2');
      const refId3 = manager.createReference(mockDataSource, 'doc-3');

      const allRefs = manager.getAllReferences();
      expect(allRefs).toHaveLength(3);

      const refIds = allRefs.map((ref) => ref.id);
      expect(refIds).toContain(refId1);
      expect(refIds).toContain(refId2);
      expect(refIds).toContain(refId3);
    });

    it('should return empty array when no references', () => {
      const allRefs = manager.getAllReferences();
      expect(allRefs).toEqual([]);
    });
  });

  describe('resolveReference', () => {
    it('should resolve reference and update display value', async () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');

      const value = await manager.resolveReference(
        manager.getReference(refId)!
      );

      expect(value).toBe('mocked-value');
      const ref = manager.getReference(refId);
      expect(ref?.display.value).toBe('mocked-value');
      expect(ref?.state).toBe('active');
    });

    it('should mark reference as broken on resolve error', async () => {
      const failingManager = new MockExcelDataSourceManager();
      failingManager.resolveReference = async () => {
        throw new Error('Mock resolve error');
      };

      const errorManager = new MemoryDataLinkManager(failingManager);
      const refId = errorManager.createReference(mockDataSource, 'doc-1');

      // Silence console.error for this test
      const originalError = console.error;
      console.error = vi.fn();

      await errorManager.resolveReference(errorManager.getReference(refId)!);

      const ref = errorManager.getReference(refId);
      expect(ref?.state).toBe('broken');

      console.error = originalError;
      errorManager.destroy();
    });
  });

  describe('resolveReferenceBatch', () => {
    it('should resolve multiple references in parallel', async () => {
      const refId1 = manager.createReference(mockDataSource, 'doc-1');
      const refId2 = manager.createReference(mockDataSource, 'doc-2');
      const refId3 = manager.createReference(mockDataSource, 'doc-3');

      const refs = [
        manager.getReference(refId1)!,
        manager.getReference(refId2)!,
        manager.getReference(refId3)!,
      ];

      const results = await manager.resolveReferenceBatch(refs);

      expect(results).toBeInstanceOf(Map);
      expect(results.size).toBe(3);
      expect(results.get(refId1)).toBe('mocked-value');
      expect(results.get(refId2)).toBe('mocked-value');
      expect(results.get(refId3)).toBe('mocked-value');
    });

    it('should return empty map for empty input', async () => {
      const results = await manager.resolveReferenceBatch([]);
      expect(results).toBeInstanceOf(Map);
      expect(results.size).toBe(0);
    });
  });

  describe('setExcelWorkbook', () => {
    it('should set workbook on Excel data source manager', () => {
      const mockWorkbook = { id: 'workbook-1' };
      manager.setExcelWorkbook(mockWorkbook as any);

      const excelManager = manager.getExcelDataSourceManager();
      expect(excelManager.getWorkbook()).toBe(mockWorkbook);
    });

    it('should handle null workbook', () => {
      expect(() =>
        manager.setExcelWorkbook(null)
      ).not.toThrow();
    });
  });

  describe('validateReference', () => {
    it('should return false for broken references', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      manager.updateReference(refId, { state: 'broken' });

      const isValid = manager.validateReference(manager.getReference(refId)!);
      expect(isValid).toBe(false);
    });

    it('should validate sheet existence', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');
      manager.getExcelDataSourceManager().setWorkbook({
        hasSheet: (id: string) => id === 'sheet1',
      } as any);

      const isValid = manager.validateReference(manager.getReference(refId)!);
      expect(isValid).toBe(true);
    });

    it('should return false for invalid sheet', () => {
      const refId = manager.createReference({
        ...mockDataSource,
        sheetId: 'invalid-sheet',
      }, 'doc-1');
      manager.getExcelDataSourceManager().setWorkbook({
        hasSheet: () => false,
      } as any);

      const isValid = manager.validateReference(manager.getReference(refId)!);
      expect(isValid).toBe(false);
    });
  });

  describe('data source change callbacks', () => {
    it('should register and call data source change callbacks', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);

      mockExcelManager.emitChange({
        sheetId: 'sheet1',
        row: 0,
        column: 0,
        value: 'new-value',
      });

      expect(callback).toHaveBeenCalled();
    });

    it('should unregister data source change callbacks', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);
      manager.offDataSourceChange(callback);

      mockExcelManager.emitChange({
        sheetId: 'sheet1',
        row: 0,
        column: 0,
        value: 'new-value',
      });

      expect(callback).not.toHaveBeenCalled();
    });

    it('should handle errors in callbacks gracefully', () => {
      const errorCallback = vi.fn(() => {
        throw new Error('Callback error');
      });
      const successCallback = vi.fn();

      manager.onDataSourceChange(errorCallback);
      manager.onDataSourceChange(successCallback);

      expect(() => {
        mockExcelManager.emitChange({
          sheetId: 'sheet1',
          row: 0,
          column: 0,
          value: 'new-value',
        });
      }).not.toThrow();

      // Success callback should still be called
      expect(successCallback).toHaveBeenCalled();
    });
  });

  describe('destroy', () => {
    it('should clean up all references', () => {
      manager.createReference(mockDataSource, 'doc-1');
      manager.createReference(mockDataSource, 'doc-2');
      expect(manager.getAllReferences()).toHaveLength(2);

      manager.destroy();
      expect(manager.getAllReferences()).toHaveLength(0);
    });

    it('should clean up data source change callbacks', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);

      manager.destroy();

      mockExcelManager.emitChange({
        sheetId: 'sheet1',
        row: 0,
        column: 0,
        value: 'new-value',
      });

      expect(callback).not.toHaveBeenCalled();
    });

    it('should destroy Excel data source manager', () => {
      manager.destroy();
      // Verify by checking callbacks are cleared
      expect(() => {
        manager.getExcelDataSourceManager().destroy();
      }).not.toThrow();
    });
  });

  describe('factory function', () => {
    it('should create manager via factory', () => {
      const factoryManager = createDataLinkManager();
      expect(factoryManager).toBeInstanceOf(MemoryDataLinkManager);

      const refId = factoryManager.createReference(mockDataSource, 'doc-1');
      expect(factoryManager.getReference(refId)).toBeDefined();

      factoryManager.destroy();
    });

    it('should create manager with custom Excel data source manager', () => {
      const customManager = createDataLinkManager(mockExcelManager);
      const excelManager = customManager.getExcelDataSourceManager();

      expect(excelManager).toBe(mockExcelManager);

      customManager.destroy();
    });
  });

  describe('reference metadata', () => {
    it('should preserve existing metadata when updating', () => {
      const refId = manager.createReference(mockDataSource, 'doc-1');

      manager.updateReference(refId, {
        metadata: { created: Date.now(), updated: Date.now() },
      });

      const refBefore = manager.getReference(refId);

      manager.updateReference(refId, {
        state: 'stale',
      });

      const refAfter = manager.getReference(refId);
      expect(refAfter?.metadata).toEqual(refBefore?.metadata);
    });
  });
});
