# DataLinkManager API

The `DataLinkManager` is responsible for managing data references between Excel data sources and Word documents.

## Overview

The `DataLinkManager` provides functionality to:
- Create, update, and delete data references
- Query references by source or document
- Resolve references to get current values from Excel
- Validate that reference sources still exist
- Monitor Excel data source changes

## Installation

```typescript
import { DataLinkManager, createDataLinkManager } from '@/core/editors/DataLinkManager';

// Create a new instance
const dataLinkManager = createDataLinkManager();
```

## Interface

```typescript
interface DataLinkManager {
  // Reference Management
  createReference(source: DataSource, target: DocumentId): ReferenceId;
  getReference(id: ReferenceId): DataReference | undefined;
  updateReference(id: ReferenceId, data: Partial<DataReference>): void;
  deleteReference(id: ReferenceId): void;

  // Reference Query
  findReferencesBySource(source: DataSource): ReferenceId[];
  findReferencesByDocument(docId: DocumentId): ReferenceId[];
  getAllReferences(): DataReference[];

  // Data Resolution
  resolveReference(ref: DataReference): Promise<any>;
  resolveReferenceBatch(refs: DataReference[]): Promise<Map<ReferenceId, any>>;

  // Excel Integration
  setExcelWorkbook(workbook: IUniverWorkbook | null): void;
  getExcelDataSourceManager(): IExcelDataSourceManager;

  // Validation
  validateReference(ref: DataReference): boolean;

  // Data Source Change Monitoring
  onDataSourceChange(callback: (event: any) => void): void;
  offDataSourceChange(callback: (event: any) => void): void;

  // Lifecycle
  destroy(): void;
}
```

**Notes:** `createReference` uses `target` as the document id; the stored reference's `target.nodeId` is set to the new reference id. `findReferencesBySource` matches by `fileId` and `sheetId` only. `resolveReference` resolves the value, updates the reference's display and state, and returns the value.

## Methods

### createReference

Creates a new data reference linking an Excel cell to a Word document.

```typescript
createReference(source: DataSource, target: DocumentId): ReferenceId
```

**Parameters:**
- `source: DataSource` - The Excel data source (file, sheet, cell/range)
- `target: DocumentId` - The target document ID

**Returns:** `ReferenceId` - The ID of the created reference

**Emits:** `Events.ReferenceCreated`

**Example:**
```typescript
const source: DataSource = {
  fileId: 'file-123',
  fileName: 'budget.xlsx',
  sheetId: 'sheet-1',
  sheetName: 'Sheet1',
  range: { startRow: 0, startCol: 0 },
  isFormula: false
};

const refId = dataLinkManager.createReference(source, 'doc-456');
console.log('Created reference:', refId);
```

---

### getReference

Retrieves a data reference by ID.

```typescript
getReference(id: ReferenceId): DataReference | undefined
```

**Parameters:**
- `id: ReferenceId` - The reference ID

**Returns:** `DataReference | undefined` - The reference, or undefined if not found

**Example:**
```typescript
const ref = dataLinkManager.getReference(refId);
if (ref) {
  console.log('Reference value:', ref.display.value);
}
```

---

### updateReference

Updates a data reference with partial data.

```typescript
updateReference(id: ReferenceId, data: Partial<DataReference>): void
```

**Parameters:**
- `id: ReferenceId` - The reference ID
- `data: Partial<DataReference>` - Partial data to update

**Emits:** `Events.ReferenceUpdated`

**Example:**
```typescript
dataLinkManager.updateReference(refId, {
  state: 'active',
  display: { value: newValue }
});
```

---

### deleteReference

Deletes a data reference by ID.

```typescript
deleteReference(id: ReferenceId): void
```

**Parameters:**
- `id: ReferenceId` - The reference ID

**Emits:** `Events.ReferenceDeleted`

**Example:**
```typescript
dataLinkManager.deleteReference(refId);
```

---

### findReferencesBySource

Finds all references that match a given data source.

```typescript
findReferencesBySource(source: DataSource): ReferenceId[]
```

**Parameters:**
- `source: DataSource` - The data source to match

**Returns:** `ReferenceId[]` - Array of matching reference IDs

**Example:**
```typescript
const source = { fileId: 'file-123', sheetId: 'sheet-1', ... };
const refs = dataLinkManager.findReferencesBySource(source);
console.log(`Found ${refs.length} references for this source`);
```

---

### findReferencesByDocument

Finds all references in a given document.

```typescript
findReferencesByDocument(docId: DocumentId): ReferenceId[]
```

**Parameters:**
- `docId: DocumentId` - The document ID

**Returns:** `ReferenceId[]` - Array of reference IDs in the document

**Example:**
```typescript
const refs = dataLinkManager.findReferencesByDocument('doc-456');
console.log(`Document has ${refs.length} references`);
```

---

### getAllReferences

Returns all references.

```typescript
getAllReferences(): DataReference[]
```

**Returns:** `DataReference[]` - Array of all references

**Example:**
```typescript
const allRefs = dataLinkManager.getAllReferences();
console.log(`Total references: ${allRefs.length}`);
```

---

### resolveReference

Resolves a reference to get the current value from Excel.

```typescript
resolveReference(ref: DataReference): Promise<any>
```

**Parameters:**
- `ref: DataReference` - The reference to resolve

**Returns:** `Promise<any>` - The resolved value

**Example:**
```typescript
const ref = dataLinkManager.getReference(refId);
if (ref) {
  const value = await dataLinkManager.resolveReference(ref);
  console.log('Resolved value:', value);
}
```

---

### resolveReferenceBatch

Resolves multiple references in parallel for better performance.

```typescript
resolveReferenceBatch(refs: DataReference[]): Promise<Map<ReferenceId, any>>
```

**Parameters:**
- `refs: DataReference[]` - The references to resolve

**Returns:** `Promise<Map<ReferenceId, any>>` - Map of reference IDs to resolved values

**Example:**
```typescript
const refs = dataLinkManager.getAllReferences();
const values = await dataLinkManager.resolveReferenceBatch(refs);
values.forEach((value, refId) => {
  console.log(`${refId}: ${value}`);
});
```

---

### setExcelWorkbook

Sets the Excel workbook for data source integration.

```typescript
setExcelWorkbook(workbook: IUniverWorkbook | null): void
```

**Parameters:**
- `workbook: IUniverWorkbook | null` - The workbook, or null to clear

**Example:**
```typescript
// Set workbook from Univer
dataLinkManager.setExcelWorkbook(univerInstance);

// Clear workbook
dataLinkManager.setExcelWorkbook(null);
```

---

### getExcelDataSourceManager

Returns the Excel data source manager.

```typescript
getExcelDataSourceManager(): IExcelDataSourceManager
```

**Returns:** `IExcelDataSourceManager` - The Excel data source manager instance

**Example:**
```typescript
const excelManager = dataLinkManager.getExcelDataSourceManager();
```

---

### validateReference

Validates that a reference's source still exists.

```typescript
validateReference(ref: DataReference): boolean
```

**Parameters:**
- `ref: DataReference` - The reference to validate

**Returns:** `boolean` - `true` if the reference source is valid

**Example:**
```typescript
const ref = dataLinkManager.getReference(refId);
if (ref && !dataLinkManager.validateReference(ref)) {
  console.warn('Reference source is invalid');
}
```

---

### onDataSourceChange

Registers a callback for Excel data source changes.

```typescript
onDataSourceChange(callback: (event: any) => void): void
```

**Parameters:**
- `callback: (event: any) => void` - The callback function

**Example:**
```typescript
dataLinkManager.onDataSourceChange((event) => {
  console.log('Data source changed:', event.sheetId);
});
```

---

### offDataSourceChange

Unregisters a data source change callback.

```typescript
offDataSourceChange(callback: (event: any) => void): void
```

**Parameters:**
- `callback: (event: any) => void` - The callback to unregister

**Example:**
```typescript
const handler = (event) => console.log(event);
dataLinkManager.onDataSourceChange(handler);
dataLinkManager.offDataSourceChange(handler);
```

---

## MemoryDataLinkManager

The default implementation stores references in memory.

```typescript
class MemoryDataLinkManager implements DataLinkManager {
  constructor(excelDataSourceManager?: IExcelDataSourceManager)
  destroy(): void
}
```

### Constructor

```typescript
constructor(excelDataSourceManager?: IExcelDataSourceManager)
```

**Parameters:**
- `excelDataSourceManager` (optional) - Custom Excel data source manager

### destroy

Cleans up resources.

```typescript
destroy(): void
```

**Example:**
```typescript
dataLinkManager.destroy();
```

## Factory Function

```typescript
function createDataLinkManager(
  excelDataSourceManager?: IExcelDataSourceManager
): DataLinkManager
```

Creates a new `DataLinkManager` instance.

**Example:**
```typescript
const dataLinkManager = createDataLinkManager();
```

## Events

The `DataLinkManager` emits the following events via the `EventBus`:

| Event | Payload | Description |
|-------|----------|-------------|
| `ReferenceCreated` | `DataReference` | Emitted when a new reference is created |
| `ReferenceUpdated` | `{ referenceId, oldState, newState }` | Emitted when a reference is updated |
| `ReferenceDeleted` | `{ referenceId }` | Emitted when a reference is deleted |

## Auto-Refresh on Data Source Changes

The `DataLinkManager` automatically monitors Excel data source changes:

```typescript
// When a cell changes in Excel:
// 1. Finds all affected references
// 2. Resolves and updates them
// 3. Emits DataChanged event
// 4. Notifies registered callbacks
```

## Usage Example: Complete Workflow

```typescript
import { createDataLinkManager } from '@/core/editors/DataLinkManager';
import { eventBus, Events } from '@/core/eventBus';

// Create the manager
const dataLinkManager = createDataLinkManager();

// Listen for events
eventBus.on(Events.ReferenceCreated, (ref) => {
  console.log('New reference created:', ref.id);
});

eventBus.on(Events.ReferenceUpdated, (data) => {
  console.log('Reference updated:', data.referenceId);
});

eventBus.on(Events.DataChanged, (changes) => {
  console.log('Data changed:', changes);
});

// Create references
const refId1 = dataLinkManager.createReference(
  { fileId: 'f1', sheetId: 's1', range: { startRow: 0, startCol: 0 } },
  'doc-1'
);

// Query references
const docRefs = dataLinkManager.findReferencesByDocument('doc-1');
console.log('Document references:', docRefs);

// Resolve references
for (const id of docRefs) {
  const ref = dataLinkManager.getReference(id);
  if (ref) {
    const value = await dataLinkManager.resolveReference(ref);
    console.log(`Resolved ${id}:`, value);
  }
}

// Update a reference
dataLinkManager.updateReference(refId1, {
  state: 'active'
});

// Delete a reference
dataLinkManager.deleteReference(refId1);

// Cleanup when done
dataLinkManager.destroy();
```
