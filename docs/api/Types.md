# Type Definitions

This document describes the core type definitions used throughout the Office Suite application.

## Import Location

```typescript
import type {
  // Basic Types
  DocumentId,
  ReferenceId,
  SnapshotId,
  CellId,

  // Data Reference Types
  ReferenceType,
  ReferenceState,
  DataSource,
  ReferenceTarget,
  ReferenceDisplay,
  ReferenceHistory,
  DataReference,

  // Document Types
  DocumentMetadata,
  WordDocument,
  NamedRange,
  ExcelDocument,

  // Snapshot Types
  DocumentSnapshot,

  // Event Types
  DataChange,
} from '@/core/types';
```

## Basic Types

### DocumentId

```typescript
type DocumentId = string;
```

Unique identifier for a document.

**Example:** `'doc-123'`, `'report-q1-2024'`

---

### ReferenceId

```typescript
type ReferenceId = string;
```

Unique identifier for a data reference.

**Example:** `'ref-abc-123'`, typically a UUID

---

### SnapshotId

```typescript
type SnapshotId = string;
```

Unique identifier for a document snapshot.

**Example:** `'snapshot-def-456'`

---

### CellId

```typescript
type CellId = string;
```

Unique identifier for a cell.

**Example:** `'sheet1!A1'`, `'Sheet1!R5C3'`

## Data Reference Types

### ReferenceType

```typescript
type ReferenceType = 'cell' | 'range' | 'named' | 'formula';
```

The type of data reference.

| Type | Description |
|------|-------------|
| `cell` | References a single cell |
| `range` | References a range of cells |
| `named` | References a named range |
| `formula` | References a formula result |

---

### ReferenceState

```typescript
type ReferenceState = 'active' | 'stale' | 'broken' | 'conflict';
```

The current state of a data reference.

| State | Description |
|--------|-------------|
| `active` | Reference is up-to-date and valid |
| `stale` | Reference may need updating (source changed recently) |
| `broken` | Reference source is no longer available |
| `conflict` | Reference has conflicting data |

---

### CellRange

```typescript
interface CellRange {
  startRow: number;
  startCol: number;
  endRow?: number;
  endCol?: number;
}
```

Describes a range of cells in a sheet.

**Properties:**
- `startRow: number` - Zero-based starting row index
- `startCol: number` - Zero-based starting column index
- `endRow?: number` - Zero-based ending row index (for ranges)
- `endCol?: number` - Zero-based ending column index (for ranges)

**Example:**
```typescript
// Single cell
const singleCell: CellRange = {
  startRow: 0,
  startCol: 0
};

// Range (A1:B2)
const rangeA1B2: CellRange = {
  startRow: 0,
  startCol: 0,
  endRow: 1,
  endCol: 1
};
```

---

### DataSource

```typescript
interface DataSource {
  fileId: string;
  fileName: string;
  sheetId: string;
  sheetName: string;
  range: CellRange;
  isFormula: boolean;
}
```

Describes the source of a data reference in Excel.

**Properties:**
- `fileId: string` - Unique identifier for the Excel file
- `fileName: string` - Display name of the file
- `sheetId: string` - Unique identifier for the sheet
- `sheetName: string` - Display name of the sheet
- `range: CellRange` - The cell or range referenced
- `isFormula: boolean` - Whether the source contains a formula

**Example:**
```typescript
const source: DataSource = {
  fileId: 'excel-123',
  fileName: 'budget.xlsx',
  sheetId: 'sheet-456',
  sheetName: 'Q1 Budget',
  range: { startRow: 0, startCol: 3 },
  isFormula: true
};
```

---

### ReferenceTarget

```typescript
interface ReferenceTarget {
  documentId: string;
  nodeId: string;
  offset?: number;
}
```

Describes where a reference is used in a document.

**Properties:**
- `documentId: string` - The target document ID
- `nodeId: string` - The node ID within the document
- `offset?: number` - Character offset for inline references

**Example:**
```typescript
const target: ReferenceTarget = {
  documentId: 'doc-789',
  nodeId: 'ref-abc-123',
  offset: 42
};
```

---

### ReferenceDisplay

```typescript
interface ReferenceDisplay {
  format: 'value' | 'expression' | 'mixed';
  value?: any;
  expression?: string;
  tooltip?: string;
}
```

Describes how a reference is displayed.

**Properties:**
- `format: 'value' | 'expression' | 'mixed'` - How to display the reference
- `value?: any` - The current value (for 'value' format)
- `expression?: string` - The formula/expression (for 'expression' format)
- `tooltip?: string` - Custom tooltip text

**Example:**
```typescript
const display: ReferenceDisplay = {
  format: 'value',
  value: 1234.56,
  tooltip: 'Q1 Total Revenue'
};

const formulaDisplay: ReferenceDisplay = {
  format: 'expression',
  value: 1234.56,
  expression: 'SUM(B2:B10)'
};
```

---

### ReferenceHistory

```typescript
interface ReferenceHistory {
  timestamp: number;
  action: 'created' | 'updated' | 'resolved' | 'broken';
  value: any;
  userId: string;
}
```

An entry in the history of a reference.

**Properties:**
- `timestamp: number` - Unix timestamp when the action occurred
- `action: 'created' | 'updated' | 'resolved' | 'broken'` - The action performed
- `value: any` - The value at this point in time
- `userId: string` - The user who performed the action

**Example:**
```typescript
const historyEntry: ReferenceHistory = {
  timestamp: Date.now(),
  action: 'updated',
  value: 1234.56,
  userId: 'user-123'
};
```

---

### DataReference

```typescript
interface DataReference {
  id: string;
  type: ReferenceType;
  source: DataSource;
  target: ReferenceTarget;
  display: ReferenceDisplay;
  state: ReferenceState;
  history: ReferenceHistory[];
  metadata?: {
    updated: number;
    created: number;
    author?: string;
  };
}
```

The complete data reference.

**Properties:**
- `id: string` - Unique reference identifier
- `type: ReferenceType` - Type of reference (cell, range, named, formula)
- `source: DataSource` - Excel data source
- `target: ReferenceTarget` - Where in the Word document
- `display: ReferenceDisplay` - Display information
- `state: ReferenceState` - Current state
- `history: ReferenceHistory[]` - Change history
- `metadata?: object` - Optional metadata

**Example:**
```typescript
const reference: DataReference = {
  id: 'ref-abc-123',
  type: 'cell',
  source: {
    fileId: 'excel-456',
    fileName: 'budget.xlsx',
    sheetId: 'sheet-789',
    sheetName: 'Q1 Budget',
    range: { startRow: 0, startCol: 3 },
    isFormula: true
  },
  target: {
    documentId: 'doc-012',
    nodeId: 'ref-abc-123'
  },
  display: {
    format: 'value',
    value: 123456.78,
    tooltip: 'Q1 Total Revenue'
  },
  state: 'active',
  history: [
    {
      timestamp: Date.now(),
      action: 'created',
      value: null,
      userId: 'system'
    }
  ],
  metadata: {
    created: Date.now(),
    updated: Date.now(),
    author: 'user-123'
  }
};
```

## Document Types

### DocumentMetadata

```typescript
interface DocumentMetadata {
  author: string;
  tags: string[];
  isTemplate: boolean;
  parentId?: string;
}
```

Metadata about a document.

**Properties:**
- `author: string` - Document author
- `tags: string[]` - Document tags
- `isTemplate: boolean` - Whether this is a template
- `parentId?: string` - Parent document ID (for templates)

---

### WordDocument

```typescript
interface WordDocument {
  id: string;
  name: string;
  content: any; // Tiptap JSON
  references: DataReference[];
  metadata: DocumentMetadata;
  version: number;
  createdAt: number;
  updatedAt: number;
}
```

A Word document with references.

**Properties:**
- `id: string` - Document identifier
- `name: string` - Document name
- `content: any` - Tiptap JSON content
- `references: DataReference[]` - Array of references
- `metadata: DocumentMetadata` - Document metadata
- `version: number` - Document version
- `createdAt: number` - Creation timestamp
- `updatedAt: number` - Last update timestamp

---

### NamedRange

```typescript
interface NamedRange {
  name: string;
  sheetId: string;
  range: CellRange;
}
```

A named range in Excel.

**Properties:**
- `name: string` - Range name
- `sheetId: string` - Sheet identifier
- `range: CellRange` - The cell range

---

### ExcelDocument

```typescript
interface ExcelDocument {
  id: string;
  name: string;
  workbook: any; // Univer JSON
  namedRanges: NamedRange[];
  metadata: DocumentMetadata;
  version: number;
  createdAt: number;
  updatedAt: number;
}
```

An Excel workbook document.

**Properties:**
- `id: string` - Document identifier
- `name: string` - Document name
- `workbook: any` - Univer workbook JSON
- `namedRanges: NamedRange[]` - Array of named ranges
- `metadata: DocumentMetadata` - Document metadata
- `version: number` - Document version
- `createdAt: number` - Creation timestamp
- `updatedAt: number` - Last update timestamp

## Snapshot Types

### DocumentSnapshot

```typescript
interface DocumentSnapshot {
  id: string;
  documentId: string;
  timestamp: number;
  data: {
    content: any;
    references: DataReference[];
  };
  checksum: string;
}
```

A point-in-time snapshot of a document.

**Properties:**
- `id: string` - Snapshot identifier
- `documentId: string` - Document identifier
- `timestamp: number` - When snapshot was created
- `data: object` - Snapshot data
  - `content: any` - Document content
  - `references: DataReference[]` - Array of references
- `checksum: string` - Checksum for change detection

## Event Types

### DataChange

```typescript
interface DataChange {
  referenceId: string;
  oldValue: any;
  newValue: any;
  timestamp: number;
  source: 'manual' | 'auto';
}
```

Describes a change in data value.

**Properties:**
- `referenceId: string` - The reference that changed
- `oldValue: any` - Previous value
- `newValue: any` - New value
- `timestamp: number` - When the change occurred
- `source: 'manual' | 'auto'` - How the change was triggered

## Type Guards

### IsReferenceActive

```typescript
function isReferenceActive(state: ReferenceState): boolean {
  return state === 'active';
}
```

### IsReferenceValid

```typescript
function isReferenceValid(state: ReferenceState): boolean {
  return state === 'active' || state === 'stale';
}
```

### IsReferenceBroken

```typescript
function isReferenceBroken(state: ReferenceState): boolean {
  return state === 'broken' || state === 'conflict';
}
```

## Helper Functions

### FormatCellReference

```typescript
function formatCellReference(row: number, col: number): string {
  const colLetter = String.fromCharCode(65 + col);
  return `${colLetter}${row + 1}`;
}
```

Converts zero-based row/col to Excel cell reference.

**Example:** `(0, 0)` -> `'A1'`, `(4, 2)` -> `'C5'`

### ParseCellReference

```typescript
function parseCellReference(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { row: 0, col: 0 };

  const col = match[1].split('').reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 65), 0);
  const row = parseInt(match[2], 10) - 1;

  return { row, col };
}
```

Converts Excel cell reference to zero-based row/col.

**Example:** `'A1'` -> `{ row: 0, col: 0 }`, `'C5'` -> `{ row: 4, col: 2 }`
