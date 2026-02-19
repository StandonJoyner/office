// Types
export * from './types';

// Event Bus
export { eventBus, Events } from './eventBus';

// Editors
export type { DataLinkManager } from './editors/DataLinkManager';
export { MemoryDataLinkManager, createDataLinkManager } from './editors/DataLinkManager';

// Excel Data Source
export type {
  IExcelDataSourceManager,
  IUniverWorkbook,
  IWorksheet,
  IRange,
  ExcelCellData,
  ExcelRangeData,
  DataSourceEvent,
} from './editors/ExcelDataSource';
export {
  ExcelDataSourceManager,
  createExcelDataSourceManager,
  DataSourceEventType,
  a1ToRowColumn,
  rowColumnToA1,
  parseCellRange,
  formatCellRange,
} from './editors/ExcelDataSource';

// Univer Workbook Adapter
export type {
  IFUniverAPI,
  IFUniverWorkbook,
  IFUniverWorksheet,
  IFUniverRange,
} from './editors/UniverWorkbookAdapter';
export {
  UniverWorkbookAdapter,
  createUniverWorkbookAdapter,
} from './editors/UniverWorkbookAdapter';

// Sync
export type { DataSyncEngine } from './sync/DataSyncEngine';
export {
  DefaultDataSyncEngine,
  createDataSyncEngine,
} from './sync/DataSyncEngine';
export type { ResolveStrategy, SyncConflict as Conflict } from './sync/DataSyncEngine';

// Traceability
export type { TraceabilityService } from './trace/TraceabilityService';
export {
  DefaultTraceabilityService,
  createTraceabilityService,
} from './trace/TraceabilityService';
export type {
  AuditLog,
  AuditActionType,
  Diff,
  DiffChange,
  AuditFilters,
  AuditAction,
  CellChange,
} from './trace/TraceabilityService';
