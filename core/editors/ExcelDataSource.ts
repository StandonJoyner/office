/**
 * Excel Data Source Manager
 *
 * Provides integration with Univer for parsing and resolving Excel cell references.
 */

import type {
  DataReference,
  DataSource,
  CellRange,
} from '../types';

/**
 * Univer Workbook interface (simplified for type safety)
 */
export interface IUniverWorkbook {
  getId(): string;
  getName(): string;
  getActiveSheet(): IWorksheet | null;
  getSheetById(sheetId: string): IWorksheet | null;
  getSheets(): IWorksheet[];
  on(type: string, callback: (...args: any[]) => void): void;
  off(type: string, callback: (...args: any[]) => void): void;
}

/**
 * Univer Worksheet interface (simplified)
 */
export interface IWorksheet {
  getSheetId(): string;
  getName(): string;
  getRange(range: string): IRange | null;
  getRangeByRow(startRow: number, startCol: number, rowCount?: number, colCount?: number): IRange | null;
  on(type: string, callback: (...args: any[]) => void): void;
  off(type: string, callback: (...args: any[]) => void): void;
}

/**
 * Univer Range interface (simplified)
 */
export interface IRange {
  getValue(): any;
  getValues(): any[][];
  getFormula(): string | null;
  getRowCount(): number;
  getColumnCount(): number;
  getRow(): number;
  getColumn(): number;
  getAddress(): string;
}

/**
 * Excel cell data
 */
export interface ExcelCellData {
  value: any;
  formula?: string | null;
  address: string;
  row: number;
  column: number;
}

/**
 * Excel range data
 */
export interface ExcelRangeData {
  values: any[][];
  rowCount: number;
  columnCount: number;
  startAddress: string;
}

/**
 * Data source event types
 */
export enum DataSourceEventType {
  CellChanged = 'cell_changed',
  RangeChanged = 'range_changed',
  SheetActivated = 'sheet_activated',
  FormulaChanged = 'formula_changed',
}

/**
 * Data source event payload
 */
export interface DataSourceEvent {
  type: DataSourceEventType;
  sheetId: string;
  row?: number;
  column?: number;
  oldValue?: any;
  newValue?: any;
  formula?: string;
}

/**
 * Excel data source manager
 *
 * Handles Excel workbook integration, cell resolution, and event monitoring.
 */
export interface IExcelDataSourceManager {
  // Workbook Management
  setWorkbook(workbook: IUniverWorkbook | null): void;
  getWorkbook(): IUniverWorkbook | null;

  // Cell/Range Resolution
  resolveCell(sheetId: string, row: number, column: number): Promise<ExcelCellData | null>;
  resolveRange(sheetId: string, range: CellRange): Promise<ExcelRangeData | null>;
  resolveReference(ref: DataReference): Promise<any>;

  // Event Monitoring
  onDataSourceChange(callback: (event: DataSourceEvent) => void): void;
  offDataSourceChange(callback: (event: DataSourceEvent) => void): void;
  enableChangeMonitoring(enabled: boolean): void;

  // Validation
  validateSheet(sheetId: string): boolean;
  validateCell(sheetId: string, row: number, column: number): boolean;

  // Cleanup
  destroy(): void;
}

/**
 * Default implementation of Excel data source manager
 */
export class ExcelDataSourceManager implements IExcelDataSourceManager {
  private workbook: IUniverWorkbook | null = null;
  private changeCallbacks = new Set<(event: DataSourceEvent) => void>();
  private monitoringEnabled = false;
  private sheetChangeListeners = new Map<string, Set<() => void>>();
  private workbookChangeListeners: (() => void)[] = [];

  /**
   * Set the active Univer workbook
   */
  setWorkbook(workbook: IUniverWorkbook): void {
    this.cleanupListeners();
    this.workbook = workbook;
    if (this.monitoringEnabled) {
      this.setupListeners();
    }
  }

  /**
   * Get the active workbook
   */
  getWorkbook(): IUniverWorkbook | null {
    return this.workbook;
  }

  /**
   * Resolve a single cell value
   */
  async resolveCell(
    sheetId: string,
    row: number,
    column: number
  ): Promise<ExcelCellData | null> {
    const workbook = this.workbook;
    if (!workbook) return null;

    const sheet = workbook.getSheetById(sheetId);
    if (!sheet) return null;

    const range = sheet.getRangeByRow(row, column, 1, 1);
    if (!range) return null;

    return {
      value: range.getValue(),
      formula: range.getFormula(),
      address: range.getAddress(),
      row: range.getRow(),
      column: range.getColumn(),
    };
  }

  /**
   * Resolve a range of cells
   */
  async resolveRange(
    sheetId: string,
    range: CellRange
  ): Promise<ExcelRangeData | null> {
    const workbook = this.workbook;
    if (!workbook) return null;

    const sheet = workbook.getSheetById(sheetId);
    if (!sheet) return null;

    const rowCount = (range.endRow ?? range.startRow) - range.startRow + 1;
    const colCount = (range.endCol ?? range.startCol) - range.startCol + 1;

    const univerRange = sheet.getRangeByRow(
      range.startRow,
      range.startCol,
      rowCount,
      colCount
    );

    if (!univerRange) return null;

    return {
      values: univerRange.getValues(),
      rowCount: univerRange.getRowCount(),
      columnCount: univerRange.getColumnCount(),
      startAddress: univerRange.getAddress(),
    };
  }

  /**
   * Resolve a data reference
   */
  async resolveReference(ref: DataReference): Promise<any> {
    const { source, type } = ref;

    // Handle single cell reference
    if (type === 'cell') {
      const cellData = await this.resolveCell(
        source.sheetId,
        source.range.startRow,
        source.range.startCol
      );
      return cellData?.value ?? null;
    }

    // Handle range reference
    if (type === 'range' && source.range.endRow !== undefined) {
      const rangeData = await this.resolveRange(source.sheetId, source.range);
      return rangeData?.values ?? null;
    }

    // Handle formula reference
    if (type === 'formula' || source.isFormula) {
      const cellData = await this.resolveCell(
        source.sheetId,
        source.range.startRow,
        source.range.startCol
      );
      return cellData?.formula ?? cellData?.value ?? null;
    }

    return null;
  }

  /**
   * Register a data source change callback
   */
  onDataSourceChange(callback: (event: DataSourceEvent) => void): void {
    this.changeCallbacks.add(callback);
    if (this.workbook && !this.monitoringEnabled) {
      this.enableChangeMonitoring(true);
    }
  }

  /**
   * Unregister a data source change callback
   */
  offDataSourceChange(callback: (event: DataSourceEvent) => void): void {
    this.changeCallbacks.delete(callback);
    if (this.changeCallbacks.size === 0) {
      this.enableChangeMonitoring(false);
    }
  }

  /**
   * Enable or disable change monitoring
   */
  enableChangeMonitoring(enabled: boolean): void {
    if (this.monitoringEnabled === enabled) return;

    this.monitoringEnabled = enabled;

    if (enabled && this.workbook) {
      this.setupListeners();
    } else {
      this.cleanupListeners();
    }
  }

  /**
   * Set up event listeners on the workbook
   */
  private setupListeners(): void {
    const workbook = this.workbook;
    if (!workbook) return;

    // Listen to workbook-level events
    const onCellChanged = (event: any) => {
      if (!event || !event.sheetId) return;

      this.notifyChange({
        type: DataSourceEventType.CellChanged,
        sheetId: event.sheetId,
        row: event.row,
        column: event.column,
        oldValue: event.oldValue,
        newValue: event.value,
      });
    };

    workbook.on('cellChange', onCellChanged);
    this.workbookChangeListeners.push(() => workbook.off('cellChange', onCellChanged));
  }

  /**
   * Clean up event listeners
   */
  private cleanupListeners(): void {
    // Remove workbook listeners
    for (const cleanup of this.workbookChangeListeners) {
      cleanup();
    }
    this.workbookChangeListeners = [];

    // Remove sheet listeners
    for (const [sheetId, listeners] of this.sheetChangeListeners) {
      const sheet = this.workbook?.getSheetById(sheetId);
      if (sheet) {
        for (const cleanup of listeners) {
          cleanup();
        }
      }
    }
    this.sheetChangeListeners.clear();
  }

  /**
   * Notify all registered callbacks of a change event
   */
  private notifyChange(event: DataSourceEvent): void {
    this.changeCallbacks.forEach((callback) => {
      try {
        callback(event);
      } catch (error) {
        console.error('Error in data source change callback:', error);
      }
    });
  }

  /**
   * Validate that a sheet exists
   */
  validateSheet(sheetId: string): boolean {
    const workbook = this.workbook;
    if (!workbook) return false;

    return workbook.getSheetById(sheetId) !== null;
  }

  /**
   * Validate that a cell exists
   */
  validateCell(sheetId: string, row: number, column: number): boolean {
    if (!this.validateSheet(sheetId)) return false;

    const sheet = this.workbook?.getSheetById(sheetId);
    if (!sheet) return false;

    const range = sheet.getRangeByRow(row, column, 1, 1);
    return range !== null;
  }

  /**
   * Clean up resources
   */
  destroy(): void {
    this.cleanupListeners();
    this.changeCallbacks.clear();
    this.workbook = null;
  }
}

/**
 * Factory function to create an Excel data source manager
 */
export function createExcelDataSourceManager(): IExcelDataSourceManager {
  return new ExcelDataSourceManager();
}

/**
 * Utility to convert A1 notation to row/column
 */
export function a1ToRowColumn(a1: string): { row: number; column: number } {
  const match = a1.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { row: 0, column: 0 };

  const column = match[1]
    .split('')
    .reduce((acc, char, index) => {
      return acc + (char.charCodeAt(0) - 64) * Math.pow(26, match[1].length - index - 1);
    }, 0);
  const row = parseInt(match[2], 10) - 1;

  return { row, column };
}

/**
 * Utility to convert row/column to A1 notation
 */
export function rowColumnToA1(row: number, column: number): string {
  const columnLetter = (() => {
    let result = '';
    let col = column + 1;
    while (col > 0) {
      col--;
      result = String.fromCharCode(65 + (col % 26)) + result;
      col = Math.floor(col / 26);
    }
    return result;
  })();

  return `${columnLetter}${row + 1}`;
}

/**
 * Utility to convert cell range to CellRange object
 */
export function parseCellRange(range: string): CellRange | null {
  // Handle single cell (e.g., "A1")
  const singleCell = range.match(/^([A-Z]+)(\d+)$/);
  if (singleCell) {
    const { row, column } = a1ToRowColumn(range);
    return { startRow: row, startCol: column };
  }

  // Handle range (e.g., "A1:B3")
  const rangeMatch = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (rangeMatch) {
    const start = a1ToRowColumn(rangeMatch[1] + rangeMatch[2]);
    const end = a1ToRowColumn(rangeMatch[3] + rangeMatch[4]);

    return {
      startRow: start.row,
      startCol: start.column,
      endRow: end.row,
      endCol: end.column,
    };
  }

  return null;
}

/**
 * Utility to format CellRange object as string
 */
export function formatCellRange(range: CellRange): string {
  const start = rowColumnToA1(range.startRow, range.startCol);

  if (range.endRow !== undefined && range.endCol !== undefined) {
    const end = rowColumnToA1(range.endRow, range.endCol);
    return `${start}:${end}`;
  }

  return start;
}
