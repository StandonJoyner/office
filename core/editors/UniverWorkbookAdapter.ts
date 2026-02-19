/**
 * Univer Workbook Adapter
 *
 * Adapts the FUniver API to the IUniverWorkbook interface for use with ExcelDataSourceManager.
 * This bridges the FUniver event system (api.addEvent) with our internal event system.
 */

import type { IUniverWorkbook, IWorksheet, IRange, DataSourceEvent, DataSourceEventType } from './ExcelDataSource';
import type { CellRange } from '../types';
import { eventBus, Events } from '../eventBus';

/**
 * FUniver API interface (simplified)
 */
export interface IFUniverAPI {
  getActiveWorkbook(): IFUniverWorkbook | null;
  addEvent(type: string, callback: (...args: any[]) => void): () => void;
}

/**
 * FUniver Workbook interface (simplified)
 */
export interface IFUniverWorkbook {
  getId(): string;
  getName(): string;
  getActiveSheet(): IFUniverWorksheet | null;
  getSheetBySheetId(sheetId: string): IFUniverWorksheet | null;
  getSheets(): IFUniverWorksheet[];
}

/**
 * FUniver Worksheet interface (simplified; matches FWorksheet getRange overloads)
 */
export interface IFUniverWorksheet {
  getSheetId(): string;
  getName(): string;
  getRange(a1: string): IFUniverRange | null;
  getRange(row: number, col: number, numRows?: number, numColumns?: number): IFUniverRange | null;
  getSelections(): any;
}

/**
 * FUniver Range interface (simplified).
 * Real FRange uses getA1Notation() for address; adapter maps it to getAddress().
 */
export interface IFUniverRange {
  getValue(): any;
  getValues(): any[][];
  setValue(value: any): void;
  getFormula(): string | null;
  getRowCount(): number;
  getColumnCount(): number;
  getRow(): number;
  getColumn(): number;
  getA1Notation?(withSheet?: boolean): string;
  getAddress?(): string;
}

/**
 * Event listener type
 */
type EventListener = (...args: any[]) => void;

/**
 * Adapter that wraps FUniver workbook and API to implement IUniverWorkbook
 */
export class UniverWorkbookAdapter implements IUniverWorkbook {
  private workbook: IFUniverWorkbook;
  private api: IFUniverAPI;
  private eventListeners = new Map<string, Set<EventListener>>();
  private disposables: (() => void)[] = [];

  constructor(workbook: IFUniverWorkbook, api: IFUniverAPI) {
    this.workbook = workbook;
    this.api = api;
    this.setupEventBridge();
  }

  /**
   * Set up event bridge between FUniver events and our event system
   */
  private setupEventBridge(): void {
    // Listen for sheet.operation events from FUniver
    const dispose = this.api.addEvent('sheet.operation' as any, (params: any) => {
      this.handleSheetOperation(params);
    });
    this.disposables.push(dispose);
  }

  /**
   * Handle sheet.operation events and translate to our event format
   */
  private handleSheetOperation(params: any): void {
    const { type, range, sheetId } = params;

    // Handle cell value changes
    if (type === 'SET_RANGE_VALUES' && range) {
      const rangeStart = (range as any).startRow;
      const colStart = (range as any).startColumn;
      const rangeValue = (range as any).getValue?.();

      // Notify our listeners
      const listeners = this.eventListeners.get('cellChange');
      if (listeners) {
        listeners.forEach((callback) => {
          try {
            callback({
              sheetId: sheetId || this.getActiveSheet()?.getSheetId(),
              row: rangeStart,
              column: colStart,
              oldValue: null,
              value: rangeValue,
            });
          } catch (error) {
            console.error('Error in cellChange listener:', error);
          }
        });
      }
    }
  }

  /**
   * Get workbook ID
   */
  getId(): string {
    return this.workbook.getId();
  }

  /**
   * Get workbook name
   */
  getName(): string {
    return this.workbook.getName();
  }

  /**
   * Get the active sheet
   */
  getActiveSheet(): IWorksheet | null {
    const sheet = this.workbook.getActiveSheet();
    return sheet ? new UniverWorksheetAdapter(sheet) : null;
  }

  /**
   * Get a sheet by ID
   */
  getSheetById(sheetId: string): IWorksheet | null {
    const sheet = this.workbook.getSheetBySheetId(sheetId);
    return sheet ? new UniverWorksheetAdapter(sheet) : null;
  }

  /**
   * Get all sheets
   */
  getSheets(): IWorksheet[] {
    return this.workbook.getSheets().map((sheet) => new UniverWorksheetAdapter(sheet));
  }

  /**
   * Register an event listener
   */
  on(type: string, callback: EventListener): void {
    if (!this.eventListeners.has(type)) {
      this.eventListeners.set(type, new Set());
    }
    this.eventListeners.get(type)!.add(callback);
  }

  /**
   * Remove an event listener
   */
  off(type: string, callback: EventListener): void {
    this.eventListeners.get(type)?.delete(callback);
  }

  /**
   * Clean up resources
   */
  destroy(): void {
    this.disposables.forEach((dispose) => dispose());
    this.disposables = [];
    this.eventListeners.clear();
  }
}

/**
 * Adapter that wraps FUniver worksheet to implement IWorksheet
 */
class UniverWorksheetAdapter implements IWorksheet {
  constructor(private worksheet: IFUniverWorksheet) {}

  getSheetId(): string {
    return this.worksheet.getSheetId();
  }

  getName(): string {
    return this.worksheet.getName();
  }

  getRange(range: string): IRange | null {
    const univerRange = this.worksheet.getRange(range);
    return univerRange ? new UniverRangeAdapter(univerRange) : null;
  }

  getRangeByRow(startRow: number, startCol: number, rowCount = 1, colCount = 1): IRange | null {
    // FUniver worksheet has getRange(row, col, numRows?, numColumns?), not getRangeByRow
    const univerRange = this.worksheet.getRange(startRow, startCol, rowCount, colCount);
    return univerRange ? new UniverRangeAdapter(univerRange) : null;
  }

  on(type: string, callback: EventListener): void {
    // Worksheet-level events are handled by the workbook adapter
    // This is a no-op as we use the workbook-level event bridge
  }

  off(type: string, callback: EventListener): void {
    // Worksheet-level events are handled by the workbook adapter
    // This is a no-op as we use the workbook-level event bridge
  }
}

/**
 * Adapter that wraps FUniver range to implement IRange
 */
class UniverRangeAdapter implements IRange {
  constructor(private range: IFUniverRange) {}

  getValue(): any {
    return this.range.getValue();
  }

  getValues(): any[][] {
    return this.range.getValues();
  }

  getFormula(): string | null {
    return this.range.getFormula();
  }

  getRowCount(): number {
    return this.range.getRowCount();
  }

  getColumnCount(): number {
    return this.range.getColumnCount();
  }

  getRow(): number {
    return this.range.getRow();
  }

  getColumn(): number {
    return this.range.getColumn();
  }

  getAddress(): string {
    // FUniver FRange exposes getA1Notation(), not getAddress()
    if (typeof (this.range as { getA1Notation?: () => string }).getA1Notation === 'function') {
      return (this.range as { getA1Notation: () => string }).getA1Notation();
    }
    if (typeof this.range.getAddress === 'function') {
      return this.range.getAddress();
    }
    return columnToA1(this.range.getColumn()) + (this.range.getRow() + 1);
  }
}

/** Convert 0-based column index to A1 column letter(s). */
function columnToA1(col: number): string {
  let s = '';
  do {
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26) - 1;
  } while (col >= 0);
  return s;
}

/**
 * Factory function to create a workbook adapter
 */
export function createUniverWorkbookAdapter(
  workbook: IFUniverWorkbook,
  api: IFUniverAPI
): IUniverWorkbook {
  return new UniverWorkbookAdapter(workbook, api);
}
