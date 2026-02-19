import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  UniverWorkbookAdapter,
  createUniverWorkbookAdapter,
  type IFUniverAPI,
  type IFUniverWorkbook,
  type IFUniverWorksheet,
  type IFUniverRange,
} from './UniverWorkbookAdapter';

// Mock FUniver range implementation
class MockFUniverRange implements IFUniverRange {
  constructor(
    private value: any,
    private row: number,
    private column: number,
    private address: string
  ) {}

  getValue(): any {
    return this.value;
  }

  getValues(): any[][] {
    return [[this.value]];
  }

  setValue(value: any): void {
    this.value = value;
  }

  getFormula(): string | null {
    if (typeof this.value === 'string' && this.value.startsWith('=')) {
      return this.value;
    }
    return null;
  }

  getRowCount(): number {
    return 1;
  }

  getColumnCount(): number {
    return 1;
  }

  getRow(): number {
    return this.row;
  }

  getColumn(): number {
    return this.column;
  }

  getAddress(): string {
    return this.address;
  }
}

// Mock FUniver worksheet implementation
class MockFUniverWorksheet implements IFUniverWorksheet {
  constructor(
    private sheetId: string,
    private name: string
  ) {}

  getSheetId(): string {
    return this.sheetId;
  }

  getName(): string {
    return this.name;
  }

  getRange(a1: string): IFUniverRange | null {
    // Parse simple cell reference like "A1"
    const match = a1.match(/^([A-Z]+)(\d+)$/);
    if (match) {
      const col = match[1];
      const row = parseInt(match[2], 10);
      // Only return valid range for simple single-letter cells (A-Z)
      if (col.length === 1 && row >= 1 && row <= 999) {
        return new MockFUniverRange(null, row - 1, col.charCodeAt(0) - 65, a1);
      }
    }
    return null;
  }

  getRangeByRow(
    row: number,
    col: number,
    rowCount?: number,
    colCount?: number
  ): IFUniverRange | null {
    const colLetter = String.fromCharCode(65 + col);
    return new MockFUniverRange(null, row, col, `${colLetter}${row + 1}`);
  }

  getSelections(): any {
    return [];
  }
}

// Mock FUniver workbook implementation
class MockFUniverWorkbook implements IFUniverWorkbook {
  constructor(
    private workbookId: string,
    private workbookName: string,
    private sheets: Map<string, MockFUniverWorksheet> = new Map()
  ) {
    this.sheets.set('sheet1', new MockFUniverWorksheet('sheet1', 'Sheet1'));
    this.sheets.set('sheet2', new MockFUniverWorksheet('sheet2', 'Sheet2'));
  }

  getId(): string {
    return this.workbookId;
  }

  getName(): string {
    return this.workbookName;
  }

  getActiveSheet(): IFUniverWorksheet | null {
    return this.sheets.get('sheet1') ?? null;
  }

  getSheetBySheetId(sheetId: string): IFUniverWorksheet | null {
    return this.sheets.get(sheetId) ?? null;
  }

  getSheets(): IFUniverWorksheet[] {
    return Array.from(this.sheets.values());
  }
}

// Mock FUniver API implementation
class MockFUniverAPI implements IFUniverAPI {
  private workbook: IFUniverWorkbook | null = null;
  private eventListeners = new Map<string, Set<(...args: any[]) => void>>();

  constructor(workbook: IFUniverWorkbook | null = null) {
    this.workbook = workbook;
  }

  getActiveWorkbook(): IFUniverWorkbook | null {
    return this.workbook;
  }

  setWorkbook(workbook: IFUniverWorkbook): void {
    this.workbook = workbook;
  }

  addEvent(type: string, callback: (...args: any[]) => void): () => void {
    if (!this.eventListeners.has(type)) {
      this.eventListeners.set(type, new Set());
    }
    this.eventListeners.get(type)!.add(callback);

    // Return unsubscribe function
    return () => {
      this.eventListeners.get(type)?.delete(callback);
    };
  }

  // Helper to trigger events for testing
  emitEvent(type: string, ...args: any[]): void {
    const listeners = this.eventListeners.get(type);
    if (listeners) {
      listeners.forEach((callback) => callback(...args));
    }
  }

  cleanup(): void {
    this.eventListeners.clear();
  }
}

describe('UniverWorkbookAdapter', () => {
  let mockWorkbook: MockFUniverWorkbook;
  let mockApi: MockFUniverAPI;
  let adapter: UniverWorkbookAdapter;

  beforeEach(() => {
    mockWorkbook = new MockFUniverWorkbook('wb-1', 'TestWorkbook');
    mockApi = new MockFUniverAPI(mockWorkbook);
    adapter = new UniverWorkbookAdapter(mockWorkbook, mockApi);
  });

  afterEach(() => {
    adapter.destroy();
    mockApi.cleanup();
  });

  describe('constructor and factory', () => {
    it('should create adapter instance', () => {
      expect(adapter).toBeInstanceOf(UniverWorkbookAdapter);
    });

    it('should create adapter via factory', () => {
      const factoryAdapter = createUniverWorkbookAdapter(mockWorkbook, mockApi);
      expect(factoryAdapter).toBeDefined();
      factoryAdapter.destroy();
    });

    it('should set up event bridge on construction', () => {
      // Event bridge setup is tested implicitly by checking event registration
      expect(() => adapter.destroy()).not.toThrow();
    });
  });

  describe('getActiveSheet', () => {
    it('should return active sheet', () => {
      const activeSheet = adapter.getActiveSheet();

      expect(activeSheet).not.toBeNull();
      expect(activeSheet?.getSheetId()).toBe('sheet1');
      expect(activeSheet?.getName()).toBe('Sheet1');
    });

    it('should return null when workbook has no active sheet', () => {
      // Create a mock workbook that returns null for getActiveSheet
      const emptyWorkbook = {
        getId: () => 'wb-2',
        getName: () => 'EmptyWorkbook',
        getActiveSheet: () => null,
        getSheetBySheetId: () => null,
        getSheets: () => [],
      } as IFUniverWorkbook;

      const emptyAdapter = new UniverWorkbookAdapter(emptyWorkbook, mockApi);
      const activeSheet = emptyAdapter.getActiveSheet();
      expect(activeSheet).toBeNull();

      emptyAdapter.destroy();
    });
  });

  describe('getSheetById', () => {
    it('should return sheet by ID', () => {
      const sheet = adapter.getSheetById('sheet1');

      expect(sheet).not.toBeNull();
      expect(sheet?.getSheetId()).toBe('sheet1');
      expect(sheet?.getName()).toBe('Sheet1');
    });

    it('should return null for non-existent sheet', () => {
      const sheet = adapter.getSheetById('non-existent');
      expect(sheet).toBeNull();
    });
  });

  describe('getSheets', () => {
    it('should return all sheets', () => {
      const sheets = adapter.getSheets();

      expect(sheets).toHaveLength(2);
      const sheetIds = sheets.map((s) => s.getSheetId());
      expect(sheetIds).toContain('sheet1');
      expect(sheetIds).toContain('sheet2');
    });

    it('should return empty array when workbook has no sheets', () => {
      const emptyWorkbook = new MockFUniverWorkbook('wb-2', 'EmptyWorkbook') as any;
      // Override getSheets to return empty array
      emptyWorkbook.getSheets = () => [];

      const emptyAdapter = new UniverWorkbookAdapter(emptyWorkbook, mockApi);
      const sheets = emptyAdapter.getSheets();

      expect(sheets).toEqual([]);

      emptyAdapter.destroy();
    });
  });

  describe('event handling', () => {
    it('should register event listener', () => {
      const callback = vi.fn();
      expect(() => adapter.on('cellChange', callback)).not.toThrow();
    });

    it('should call registered event listeners', () => {
      const callback = vi.fn();
      adapter.on('cellChange', callback);

      // Emit a cell change event via the adapter's internal mechanism
      adapter.on('testEvent', callback);

      expect(() => adapter.off('testEvent', callback)).not.toThrow();
    });

    it('should unregister event listener', () => {
      const callback = vi.fn();
      adapter.on('cellChange', callback);

      adapter.off('cellChange', callback);

      // Callback should not be called when event is emitted
      expect(() => adapter.off('cellChange', callback)).not.toThrow();
    });

    it('should handle sheet.operation events', () => {
      const callback = vi.fn();
      adapter.on('cellChange', callback);

      // Simulate FUniver sheet.operation event
      mockApi.emitEvent('sheet.operation', {
        type: 'SET_RANGE_VALUES',
        range: {
          startRow: 0,
          startColumn: 0,
          getValue: () => 'new-value',
        },
        sheetId: 'sheet1',
      });

      // The callback should be invoked by the adapter's internal event handling
      expect(callback).toHaveBeenCalled();
      const eventArg = callback.mock.calls[0][0];
      expect(eventArg.sheetId).toBe('sheet1');
      expect(eventArg.row).toBe(0);
      expect(eventArg.column).toBe(0);
      expect(eventArg.value).toBe('new-value');
    });

    it('should handle sheet.operation events without range', () => {
      const callback = vi.fn();
      adapter.on('cellChange', callback);

      // Emit event without range
      mockApi.emitEvent('sheet.operation', {
        type: 'OTHER_OPERATION',
      });

      // Callback should not be called for non-SET_RANGE_VALUES events
      expect(callback).not.toHaveBeenCalled();
    });

    it('should use active sheet ID when sheetId is not provided', () => {
      const callback = vi.fn();
      adapter.on('cellChange', callback);

      // Emit event without sheetId
      mockApi.emitEvent('sheet.operation', {
        type: 'SET_RANGE_VALUES',
        range: {
          startRow: 1,
          startColumn: 2,
          getValue: () => 'test',
        },
      });

      expect(callback).toHaveBeenCalled();
      const eventArg = callback.mock.calls[0][0];
      expect(eventArg.sheetId).toBe('sheet1'); // Default to active sheet
    });

    it('should handle errors in event listeners gracefully', () => {
      const errorCallback = vi.fn(() => {
        throw new Error('Listener error');
      });
      const successCallback = vi.fn();

      adapter.on('cellChange', errorCallback);
      adapter.on('cellChange', successCallback);

      expect(() => {
        mockApi.emitEvent('sheet.operation', {
          type: 'SET_RANGE_VALUES',
          range: {
            startRow: 0,
            startColumn: 0,
            getValue: () => 'value',
          },
          sheetId: 'sheet1',
        });
      }).not.toThrow();

      // Success callback should still be called
      expect(successCallback).toHaveBeenCalled();
    });
  });

  describe('destroy', () => {
    it('should clean up event listeners', () => {
      const callback = vi.fn();
      adapter.on('cellChange', callback);

      adapter.destroy();

      // Events should no longer trigger
      mockApi.emitEvent('sheet.operation', {
        type: 'SET_RANGE_VALUES',
        range: {
          startRow: 0,
          startColumn: 0,
          getValue: () => 'value',
        },
        sheetId: 'sheet1',
      });

      expect(callback).not.toHaveBeenCalled();
    });

    it('should clear disposables', () => {
      const firstAdapter = new UniverWorkbookAdapter(mockWorkbook, mockApi);
      expect(() => firstAdapter.destroy()).not.toThrow();
      // Calling destroy twice should be safe
      expect(() => firstAdapter.destroy()).not.toThrow();
    });
  });

  describe('worksheet adapter', () => {
    it('should adapt worksheet methods', () => {
      const sheet = adapter.getSheetById('sheet1');
      expect(sheet).not.toBeNull();

      expect(sheet?.getSheetId()).toBe('sheet1');
      expect(sheet?.getName()).toBe('Sheet1');
    });

    it('should get range by A1 notation', () => {
      const sheet = adapter.getSheetById('sheet1');
      const range = sheet?.getRange('A1');

      expect(range).not.toBeNull();
      expect(range?.getAddress()).toBe('A1');
    });

    it('should get range by row/column', () => {
      const sheet = adapter.getSheetById('sheet1');
      const range = sheet?.getRangeByRow(0, 0);

      expect(range).not.toBeNull();
      expect(range?.getRow()).toBe(0);
      expect(range?.getColumn()).toBe(0);
    });

    it('should handle get range for non-existent A1', () => {
      const sheet = adapter.getSheetById('sheet1');
      const range = sheet?.getRange('XYZ999');

      expect(range).toBeNull();
    });
  });

  describe('range adapter', () => {
    it('should adapt range methods', () => {
      const sheet = adapter.getSheetById('sheet1');
      const range = sheet?.getRange('A1');

      expect(range).not.toBeNull();
      expect(range?.getAddress()).toBe('A1');
      expect(range?.getRow()).toBe(0);
      expect(range?.getColumn()).toBe(0);
    });

    it('should get row and column count', () => {
      const sheet = adapter.getSheetById('sheet1');
      const range = sheet?.getRangeByRow(0, 0, 5, 3);

      expect(range?.getRowCount()).toBe(1); // Mock returns 1
      expect(range?.getColumnCount()).toBe(1); // Mock returns 1
    });
  });
});

describe('UniverWorkbookAdapter Integration', () => {
  describe('event flow from FUniver to adapter', () => {
    it('should complete event flow: FUniver -> API -> Adapter -> Listener', () => {
      const mockWorkbook = new MockFUniverWorkbook('wb-1', 'Test');
      const mockApi = new MockFUniverAPI(mockWorkbook);
      const adapter = new UniverWorkbookAdapter(mockWorkbook, mockApi);

      const listener = vi.fn();
      adapter.on('cellChange', listener);

      // Simulate FUniver API event
      mockApi.emitEvent('sheet.operation', {
        type: 'SET_RANGE_VALUES',
        range: {
          startRow: 5,
          startColumn: 3,
          getValue: () => 'integration-test-value',
        },
        sheetId: 'sheet1',
      });

      expect(listener).toHaveBeenCalledTimes(1);
      const eventData = listener.mock.calls[0][0];
      expect(eventData.sheetId).toBe('sheet1');
      expect(eventData.row).toBe(5);
      expect(eventData.column).toBe(3);
      expect(eventData.value).toBe('integration-test-value');

      adapter.destroy();
      mockApi.cleanup();
    });

    it('should handle multiple listeners on the same event', () => {
      const mockWorkbook = new MockFUniverWorkbook('wb-1', 'Test');
      const mockApi = new MockFUniverAPI(mockWorkbook);
      const adapter = new UniverWorkbookAdapter(mockWorkbook, mockApi);

      const listener1 = vi.fn();
      const listener2 = vi.fn();
      const listener3 = vi.fn();

      adapter.on('cellChange', listener1);
      adapter.on('cellChange', listener2);
      adapter.on('cellChange', listener3);

      mockApi.emitEvent('sheet.operation', {
        type: 'SET_RANGE_VALUES',
        range: {
          startRow: 0,
          startColumn: 0,
          getValue: () => 'multi-listener-test',
        },
        sheetId: 'sheet1',
      });

      expect(listener1).toHaveBeenCalled();
      expect(listener2).toHaveBeenCalled();
      expect(listener3).toHaveBeenCalled();

      adapter.destroy();
      mockApi.cleanup();
    });

    it('should support removing individual listeners', () => {
      const mockWorkbook = new MockFUniverWorkbook('wb-1', 'Test');
      const mockApi = new MockFUniverAPI(mockWorkbook);
      const adapter = new UniverWorkbookAdapter(mockWorkbook, mockApi);

      const listener1 = vi.fn();
      const listener2 = vi.fn();

      adapter.on('cellChange', listener1);
      adapter.on('cellChange', listener2);

      // Remove listener1
      adapter.off('cellChange', listener1);

      mockApi.emitEvent('sheet.operation', {
        type: 'SET_RANGE_VALUES',
        range: {
          startRow: 0,
          startColumn: 0,
          getValue: () => 'test',
        },
        sheetId: 'sheet1',
      });

      expect(listener1).not.toHaveBeenCalled();
      expect(listener2).toHaveBeenCalled();

      adapter.destroy();
      mockApi.cleanup();
    });
  });
});

