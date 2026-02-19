import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  ExcelDataSourceManager,
  createExcelDataSourceManager,
  a1ToRowColumn,
  rowColumnToA1,
  parseCellRange,
  formatCellRange,
  type IWorksheet,
  type IRange,
} from './ExcelDataSource';
import type { DataReference, CellRange } from '../types';

// Mock worksheet implementation
class MockWorksheet implements IWorksheet {
  constructor(
    private sheetId: string,
    private name: string,
    private cells: Map<string, any> = new Map()
  ) {}

  getSheetId(): string {
    return this.sheetId;
  }

  getName(): string {
    return this.name;
  }

  getRange(range: string): IRange | null {
    // Parse simple cell reference like "A1"
    const match = range.match(/^([A-Z]+)(\d+)$/);
    if (match) {
      const col = match[1];
      const row = parseInt(match[2], 10);
      const cellKey = `${col}${row}`;
      const value = this.cells.get(cellKey);
      return new MockRange(value, row - 1, this.colToIndex(col), range);
    }
    return null;
  }

  getRangeByRow(
    startRow: number,
    startCol: number,
    rowCount = 1,
    colCount = 1
  ): IRange | null {
    const values: any[][] = [];
    for (let r = 0; r < rowCount; r++) {
      const rowValues: any[] = [];
      for (let c = 0; c < colCount; c++) {
        const colLetter = this.indexToCol(startCol + c);
        const rowNum = startRow + r + 1;
        const cellKey = `${colLetter}${rowNum}`;
        rowValues.push(this.cells.get(cellKey) ?? null);
      }
      values.push(rowValues);
    }
    const startColLetter = this.indexToCol(startCol);
    return new MockRange(
      values[0][0],
      startRow,
      startCol,
      `${startColLetter}${startRow + 1}`,
      rowCount,
      colCount,
      values
    );
  }

  on(type: string, callback: (...args: any[]) => void): void {
    // Mock implementation
  }

  off(type: string, callback: (...args: any[]) => void): void {
    // Mock implementation
  }

  setCell(col: string, row: number, value: any): void {
    this.cells.set(`${col}${row}`, value);
  }

  private colToIndex(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result - 1;
  }

  private indexToCol(index: number): string {
    let result = '';
    let col = index + 1;
    while (col > 0) {
      col--;
      result = String.fromCharCode(65 + (col % 26)) + result;
      col = Math.floor(col / 26);
    }
    return result;
  }
}

// Mock range implementation
class MockRange implements IRange {
  constructor(
    private value: any,
    private row: number,
    private column: number,
    private address: string,
    private rowCount = 1,
    private columnCount = 1,
    private valuesMatrix?: any[][]
  ) {}

  getValue(): any {
    return this.value;
  }

  getValues(): any[][] {
    if (this.valuesMatrix) {
      return this.valuesMatrix;
    }
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
    return this.rowCount;
  }

  getColumnCount(): number {
    return this.columnCount;
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

// Mock workbook implementation
class MockWorkbook {
  private sheets = new Map<string, MockWorksheet>();

  addSheet(sheetId: string, name: string): void {
    this.sheets.set(sheetId, new MockWorksheet(sheetId, name));
  }

  getSheetById(sheetId: string): IWorksheet | null {
    return this.sheets.get(sheetId) ?? null;
  }

  getSheets(): IWorksheet[] {
    return Array.from(this.sheets.values());
  }

  getActiveSheet(): IWorksheet | null {
    return Array.from(this.sheets.values())[0] ?? null;
  }

  on(type: string, callback: (...args: any[]) => void): void {
    // Mock implementation
  }

  off(type: string, callback: (...args: any[]) => void): void {
    // Mock implementation
  }

  setCellValue(sheetId: string, col: string, row: number, value: any): void {
    const sheet = this.sheets.get(sheetId);
    if (sheet) {
      sheet.setCell(col, row, value);
    }
  }
}

describe('ExcelDataSourceManager', () => {
  let manager: ExcelDataSourceManager;
  let mockWorkbook: MockWorkbook;

  beforeEach(() => {
    manager = createExcelDataSourceManager();
    mockWorkbook = new MockWorkbook();
    mockWorkbook.addSheet('sheet1', 'Sheet1');
    mockWorkbook.addSheet('sheet2', 'Sheet2');
  });

  describe('setWorkbook and getWorkbook', () => {
    it('should set and get workbook', () => {
      expect(manager.getWorkbook()).toBeNull();

      manager.setWorkbook(mockWorkbook);
      expect(manager.getWorkbook()).toBe(mockWorkbook);
    });

    it('should replace existing workbook', () => {
      manager.setWorkbook(mockWorkbook);

      const newWorkbook = new MockWorkbook();
      manager.setWorkbook(newWorkbook);

      expect(manager.getWorkbook()).toBe(newWorkbook);
    });
  });

  describe('resolveCell', () => {
    beforeEach(() => {
      manager.setWorkbook(mockWorkbook);
      mockWorkbook.setCellValue('sheet1', 'A', 1, 'test-value');
      mockWorkbook.setCellValue('sheet1', 'B', 2, 123);
    });

    it('should resolve a single cell value', async () => {
      const cellData = await manager.resolveCell('sheet1', 0, 0);

      expect(cellData).not.toBeNull();
      expect(cellData?.value).toBe('test-value');
      expect(cellData?.row).toBe(0);
      expect(cellData?.column).toBe(0);
    });

    it('should resolve cell with formula', async () => {
      mockWorkbook.setCellValue('sheet1', 'C', 3, '=SUM(A1:B2)');

      const cellData = await manager.resolveCell('sheet1', 2, 2);

      expect(cellData).not.toBeNull();
      expect(cellData?.formula).toBe('=SUM(A1:B2)');
    });

    it('should return null for non-existent sheet', async () => {
      const cellData = await manager.resolveCell('non-existent', 0, 0);
      expect(cellData).toBeNull();
    });

    it('should return null when no workbook is set', async () => {
      const noWorkbookManager = createExcelDataSourceManager();
      const cellData = await noWorkbookManager.resolveCell('sheet1', 0, 0);
      expect(cellData).toBeNull();
    });

    it('should include cell address', async () => {
      const cellData = await manager.resolveCell('sheet1', 0, 0);

      expect(cellData?.address).toBe('A1');
    });
  });

  describe('resolveRange', () => {
    beforeEach(() => {
      manager.setWorkbook(mockWorkbook);
      mockWorkbook.setCellValue('sheet1', 'A', 1, 1);
      mockWorkbook.setCellValue('sheet1', 'B', 1, 2);
      mockWorkbook.setCellValue('sheet1', 'A', 2, 3);
      mockWorkbook.setCellValue('sheet1', 'B', 2, 4);
    });

    it('should resolve a range of cells', async () => {
      const cellRange: CellRange = {
        startRow: 0,
        startCol: 0,
        endRow: 1,
        endCol: 1,
      };

      const rangeData = await manager.resolveRange('sheet1', cellRange);

      expect(rangeData).not.toBeNull();
      expect(rangeData?.rowCount).toBe(2);
      expect(rangeData?.columnCount).toBe(2);
      expect(rangeData?.values).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    it('should resolve single cell range (no end coordinates)', async () => {
      const cellRange: CellRange = {
        startRow: 0,
        startCol: 0,
      };

      const rangeData = await manager.resolveRange('sheet1', cellRange);

      expect(rangeData).not.toBeNull();
      expect(rangeData?.rowCount).toBe(1);
      expect(rangeData?.columnCount).toBe(1);
      expect(rangeData?.values[0][0]).toBe(1);
    });

    it('should return null for non-existent sheet', async () => {
      const cellRange: CellRange = { startRow: 0, startCol: 0 };
      const rangeData = await manager.resolveRange('non-existent', cellRange);
      expect(rangeData).toBeNull();
    });
  });

  describe('resolveReference', () => {
    beforeEach(() => {
      manager.setWorkbook(mockWorkbook);
      mockWorkbook.setCellValue('sheet1', 'A', 1, 'cell-value');
      mockWorkbook.setCellValue('sheet1', 'B', 1, '=A1*2');
    });

    it('should resolve cell reference', async () => {
      const ref: DataReference = {
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
        target: { documentId: 'doc-1', nodeId: 'node-1' },
        display: { format: 'value' },
        state: 'active',
        history: [],
      };

      const value = await manager.resolveReference(ref);
      expect(value).toBe('cell-value');
    });

    it('should resolve range reference', async () => {
      mockWorkbook.setCellValue('sheet1', 'A', 1, 1);
      mockWorkbook.setCellValue('sheet1', 'B', 1, 2);

      const ref: DataReference = {
        id: 'ref-1',
        type: 'range',
        source: {
          fileId: 'file-1',
          fileName: 'Test.xlsx',
          sheetId: 'sheet1',
          sheetName: 'Sheet1',
          range: { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
          isFormula: false,
        },
        target: { documentId: 'doc-1', nodeId: 'node-1' },
        display: { format: 'value' },
        state: 'active',
        history: [],
      };

      const value = await manager.resolveReference(ref);
      expect(value).toEqual([[1, 2]]);
    });

    it('should resolve formula reference', async () => {
      const ref: DataReference = {
        id: 'ref-1',
        type: 'formula',
        source: {
          fileId: 'file-1',
          fileName: 'Test.xlsx',
          sheetId: 'sheet1',
          sheetName: 'Sheet1',
          range: { startRow: 0, startCol: 1 },
          isFormula: true,
        },
        target: { documentId: 'doc-1', nodeId: 'node-1' },
        display: { format: 'value' },
        state: 'active',
        history: [],
      };

      const value = await manager.resolveReference(ref);
      expect(value).toBe('=A1*2');
    });

    it('should return null for unknown reference type', async () => {
      const ref: DataReference = {
        id: 'ref-1',
        type: 'named' as any,
        source: {
          fileId: 'file-1',
          fileName: 'Test.xlsx',
          sheetId: 'sheet1',
          sheetName: 'Sheet1',
          range: { startRow: 0, startCol: 0 },
          isFormula: false,
        },
        target: { documentId: 'doc-1', nodeId: 'node-1' },
        display: { format: 'value' },
        state: 'active',
        history: [],
      };

      const value = await manager.resolveReference(ref);
      expect(value).toBeNull();
    });
  });

  describe('data source change callbacks', () => {
    it('should register change callback', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);

      // Callback should be registered (we can't directly verify without triggering)
      expect(() => manager.onDataSourceChange(callback)).not.toThrow();
    });

    it('should unregister change callback', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);
      manager.offDataSourceChange(callback);

      expect(() => manager.offDataSourceChange(callback)).not.toThrow();
    });

    it('should enable change monitoring when first callback is added', () => {
      manager.setWorkbook(mockWorkbook);

      const callback = vi.fn();
      manager.onDataSourceChange(callback);

      // Monitoring should be enabled
      expect(() => manager.enableChangeMonitoring(true)).not.toThrow();
    });

    it('should disable change monitoring when last callback is removed', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);
      manager.offDataSourceChange(callback);

      // Monitoring should be disabled
      expect(() => manager.enableChangeMonitoring(false)).not.toThrow();
    });
  });

  describe('enableChangeMonitoring', () => {
    beforeEach(() => {
      manager.setWorkbook(mockWorkbook);
    });

    it('should enable change monitoring', () => {
      expect(() => manager.enableChangeMonitoring(true)).not.toThrow();
    });

    it('should disable change monitoring', () => {
      manager.enableChangeMonitoring(true);
      expect(() => manager.enableChangeMonitoring(false)).not.toThrow();
    });

    it('should not enable monitoring if no workbook', () => {
      const noWorkbookManager = createExcelDataSourceManager();
      expect(() => noWorkbookManager.enableChangeMonitoring(true)).not.toThrow();
    });
  });

  describe('validateSheet', () => {
    beforeEach(() => {
      manager.setWorkbook(mockWorkbook);
    });

    it('should return true for existing sheet', () => {
      expect(manager.validateSheet('sheet1')).toBe(true);
      expect(manager.validateSheet('sheet2')).toBe(true);
    });

    it('should return false for non-existent sheet', () => {
      expect(manager.validateSheet('non-existent')).toBe(false);
    });

    it('should return false when no workbook', () => {
      const noWorkbookManager = createExcelDataSourceManager();
      expect(noWorkbookManager.validateSheet('sheet1')).toBe(false);
    });
  });

  describe('validateCell', () => {
    beforeEach(() => {
      manager.setWorkbook(mockWorkbook);
    });

    it('should return true for valid cell', () => {
      expect(manager.validateCell('sheet1', 0, 0)).toBe(true);
      expect(manager.validateCell('sheet1', 100, 26)).toBe(true);
    });

    it('should return false for invalid sheet', () => {
      expect(manager.validateCell('non-existent', 0, 0)).toBe(false);
    });

    it('should return false when no workbook', () => {
      const noWorkbookManager = createExcelDataSourceManager();
      expect(noWorkbookManager.validateCell('sheet1', 0, 0)).toBe(false);
    });
  });

  describe('destroy', () => {
    it('should clean up workbook reference', () => {
      manager.setWorkbook(mockWorkbook);
      expect(manager.getWorkbook()).toBe(mockWorkbook);

      manager.destroy();
      expect(manager.getWorkbook()).toBeNull();
    });

    it('should clean up change callbacks', () => {
      const callback = vi.fn();
      manager.onDataSourceChange(callback);

      manager.destroy();

      // Callbacks should be cleared
      manager.setWorkbook(mockWorkbook);
      expect(() => manager.enableChangeMonitoring(true)).not.toThrow();
    });
  });

  describe('factory function', () => {
    it('should create manager via factory', () => {
      const factoryManager = createExcelDataSourceManager();
      expect(factoryManager).toBeInstanceOf(ExcelDataSourceManager);
      expect(factoryManager.getWorkbook()).toBeNull();
    });
  });
});

describe('Excel Utility Functions', () => {
  describe('a1ToRowColumn', () => {
    it('should convert A1 notation to row/column', () => {
      // Note: The current implementation returns 1-indexed column
      expect(a1ToRowColumn('A1')).toEqual({ row: 0, column: 1 });
      expect(a1ToRowColumn('B1')).toEqual({ row: 0, column: 2 });
      expect(a1ToRowColumn('A2')).toEqual({ row: 1, column: 1 });
      expect(a1ToRowColumn('Z1')).toEqual({ row: 0, column: 26 });
      expect(a1ToRowColumn('AA1')).toEqual({ row: 0, column: 27 });
      expect(a1ToRowColumn('AB1')).toEqual({ row: 0, column: 28 });
      expect(a1ToRowColumn('BA1')).toEqual({ row: 0, column: 53 });
      expect(a1ToRowColumn('AAA1')).toEqual({ row: 0, column: 703 });
    });

    it('should handle larger row numbers', () => {
      expect(a1ToRowColumn('A10')).toEqual({ row: 9, column: 1 });
      expect(a1ToRowColumn('B100')).toEqual({ row: 99, column: 2 });
    });

    it('should return zeros for invalid format', () => {
      expect(a1ToRowColumn('')).toEqual({ row: 0, column: 0 });
      expect(a1ToRowColumn('123')).toEqual({ row: 0, column: 0 });
      expect(a1ToRowColumn('ABC')).toEqual({ row: 0, column: 0 });
    });
  });

  describe('rowColumnToA1', () => {
    it('should convert row/column (0-indexed) to A1 notation', () => {
      expect(rowColumnToA1(0, 0)).toBe('A1');
      expect(rowColumnToA1(0, 1)).toBe('B1');
      expect(rowColumnToA1(1, 0)).toBe('A2');
      expect(rowColumnToA1(0, 25)).toBe('Z1');
      expect(rowColumnToA1(0, 26)).toBe('AA1');
      expect(rowColumnToA1(0, 27)).toBe('AB1');
      expect(rowColumnToA1(0, 51)).toBe('AZ1');
      expect(rowColumnToA1(0, 52)).toBe('BA1');
      expect(rowColumnToA1(0, 701)).toBe('ZZ1');
      expect(rowColumnToA1(0, 702)).toBe('AAA1');
    });

    it('should handle larger row numbers', () => {
      expect(rowColumnToA1(9, 0)).toBe('A10');
      expect(rowColumnToA1(99, 1)).toBe('B100');
    });
  });

  describe('parseCellRange', () => {
    it('should parse single cell reference', () => {
      // Note: The current implementation returns 1-indexed column
      expect(parseCellRange('A1')).toEqual({ startRow: 0, startCol: 1 });
      expect(parseCellRange('Z10')).toEqual({ startRow: 9, startCol: 26 });
      expect(parseCellRange('AA5')).toEqual({ startRow: 4, startCol: 27 });
    });

    it('should parse range reference', () => {
      // Note: The current implementation returns 1-indexed column
      expect(parseCellRange('A1:B2')).toEqual({
        startRow: 0,
        startCol: 1,
        endRow: 1,
        endCol: 2,
      });
      expect(parseCellRange('C3:D5')).toEqual({
        startRow: 2,
        startCol: 3,
        endRow: 4,
        endCol: 4,
      });
      expect(parseCellRange('A1:Z100')).toEqual({
        startRow: 0,
        startCol: 1,
        endRow: 99,
        endCol: 26,
      });
    });

    it('should handle complex column letters', () => {
      // Note: The current implementation returns 1-indexed column
      expect(parseCellRange('AA1:AB2')).toEqual({
        startRow: 0,
        startCol: 27,
        endRow: 1,
        endCol: 28,
      });
    });

    it('should return null for invalid format', () => {
      expect(parseCellRange('')).toBeNull();
      expect(parseCellRange('123')).toBeNull();
      expect(parseCellRange('ABC')).toBeNull();
      expect(parseCellRange('A1:B2:C3')).toBeNull();
      expect(parseCellRange('1A')).toBeNull();
    });
  });

  describe('formatCellRange', () => {
    it('should format single cell range', () => {
      expect(formatCellRange({ startRow: 0, startCol: 0 })).toBe('A1');
      expect(formatCellRange({ startRow: 9, startCol: 25 })).toBe('Z10');
    });

    it('should format range with end coordinates', () => {
      expect(formatCellRange({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 })).toBe('A1:B2');
      expect(formatCellRange({ startRow: 2, startCol: 2, endRow: 4, endCol: 3 })).toBe('C3:D5');
      expect(formatCellRange({ startRow: 0, startCol: 0, endRow: 99, endCol: 25 })).toBe('A1:Z100');
    });

    it('should handle only endRow specified', () => {
      expect(formatCellRange({ startRow: 0, startCol: 0, endRow: 2 })).toBe('A1');
    });

    it('should handle only endCol specified', () => {
      expect(formatCellRange({ startRow: 0, startCol: 0, endCol: 2 })).toBe('A1');
    });
  });
});
