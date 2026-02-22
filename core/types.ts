// Basic Types
export type DocumentId = string;
export type ReferenceId = string;
export type SnapshotId = string;
export type CellId = string;

// Data Reference Types
export type ReferenceType = 'cell' | 'range' | 'named' | 'formula';
export type ReferenceState = 'active' | 'stale' | 'broken' | 'conflict';

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow?: number;
  endCol?: number;
}

export interface DataSource {
  fileId: string;
  fileName: string;
  sheetId: string;
  sheetName: string;
  range: CellRange;
  isFormula: boolean;
}

export interface ReferenceTarget {
  documentId: string;
  nodeId: string;
  offset?: number;
}

export interface ReferenceDisplay {
  format: 'value' | 'expression' | 'mixed';
  value?: any;
  expression?: string;
  tooltip?: string;
  tableMeta?: RangeTableMeta;
}

export interface ReferenceHistory {
  timestamp: number;
  action: 'created' | 'updated' | 'resolved' | 'broken';
  value: any;
  userId: string;
}

export interface RangeTableMeta {
  rowCount: number;
  colCount: number;
  syncMode: 'manual' | 'auto';
  preserveFormatting: boolean;
}

export interface DataReference {
  id: string;
  type: ReferenceType;
  source: DataSource;
  target: ReferenceTarget;
  display: ReferenceDisplay;
  state: ReferenceState;
  history: ReferenceHistory[];
  metadata?: {
    updated: number;
  created?: number;
  author?: string;
  };
}

// Document Types
export interface DocumentMetadata {
  author: string;
  tags: string[];
  isTemplate: boolean;
  parentId?: string;
}

export interface WordDocument {
  id: string;
  name: string;
  content: any; // Tiptap JSON
  references: DataReference[];
  metadata: DocumentMetadata;
  version: number;
  createdAt: number;
  updatedAt: number;
}

export interface NamedRange {
  name: string;
  sheetId: string;
  range: CellRange;
}

export interface ExcelDocument {
  id: string;
  name: string;
  workbook: any; // Univer JSON
  namedRanges: NamedRange[];
  metadata: DocumentMetadata;
  version: number;
  createdAt: number;
  updatedAt: number;
}

// Snapshot Types
export interface DocumentSnapshot {
  id: string;
  documentId: string;
  timestamp: number;
  data: {
    content: any;
    references: DataReference[];
  };
  checksum: string;
}

// Event Types
export interface DataChange {
  referenceId: string;
  oldValue: any;
  newValue: any;
  timestamp: number;
  source: 'manual' | 'auto';
}
