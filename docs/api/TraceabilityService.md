# TraceabilityService API

The `TraceabilityService` provides complete data traceability capabilities including version snapshots, change history, dependency analysis, and audit logging.

## Overview

The `TraceabilityService` provides functionality to:
- Create and manage document snapshots
- Compare snapshots to detect changes
- Track reference and cell history
- Analyze dependencies between documents
- Log and query audit actions

## Installation

```typescript
import {
  TraceabilityService,
  createTraceabilityService
} from '@/core/trace/TraceabilityService';

// Create a new instance
const traceabilityService = createTraceabilityService(
  (docId) => getDocumentContent(docId),
  (docId) => getDocumentReferences(docId)
);
```

## Interface

```typescript
interface TraceabilityService {
  // Snapshot Management
  createSnapshot(docId: DocumentId): Promise<SnapshotId>;
  getSnapshot(id: SnapshotId): DocumentSnapshot | null;
  listSnapshots(docId: DocumentId): DocumentSnapshot[];
  deleteSnapshot(id: SnapshotId): void;
  compareSnapshots(id1: SnapshotId, id2: SnapshotId): Diff | null;

  // Change History
  getReferenceHistory(refId: ReferenceId): DataReference['history'];
  getCellHistory(cellId: CellId): CellChange[];
  updateReferenceHistory(refId: ReferenceId, history: DataReference['history']): void;

  // Dependency Analysis
  getDependents(cellId: CellId): ReferenceId[];
  getDependencies(docId: DocumentId): DataSource[];

  // Audit Logging
  logAction(action: AuditAction): void;
  getAuditLog(filters?: AuditFilters): AuditLog[];
  clearAuditLog(): void;
}
```

## Methods

### createSnapshot

Creates a snapshot of a document with its references.

```typescript
createSnapshot(docId: DocumentId): Promise<SnapshotId>
```

**Parameters:**
- `docId: DocumentId` - The document ID to snapshot

**Returns:** `Promise<SnapshotId>` - The ID of the created snapshot

**Emits:** `Events.ReferenceCreated` with snapshot info

**Example:**
```typescript
const snapshotId = await traceabilityService.createSnapshot('doc-456');
console.log('Snapshot created:', snapshotId);
```

---

### getSnapshot

Retrieves a specific snapshot by ID.

```typescript
getSnapshot(id: SnapshotId): DocumentSnapshot | null
```

**Parameters:**
- `id: SnapshotId` - The snapshot ID

**Returns:** `DocumentSnapshot | null` - The snapshot, or null if not found

**Example:**
```typescript
const snapshot = traceabilityService.getSnapshot(snapshotId);
if (snapshot) {
  console.log('Snapshot from:', new Date(snapshot.timestamp));
}
```

---

### listSnapshots

Lists all snapshots for a document, sorted by timestamp.

```typescript
listSnapshots(docId: DocumentId): DocumentSnapshot[]
```

**Parameters:**
- `docId: DocumentId` - The document ID

**Returns:** `DocumentSnapshot[]` - Array of snapshots, oldest first

**Example:**
```typescript
const snapshots = traceabilityService.listSnapshots('doc-456');
console.log(`Found ${snapshots.length} snapshots`);
snapshots.forEach((snapshot, index) => {
  console.log(`${index + 1}. ${new Date(snapshot.timestamp).toISOString()}`);
});
```

---

### deleteSnapshot

Deletes a snapshot by ID.

```typescript
deleteSnapshot(id: SnapshotId): void
```

**Parameters:**
- `id: SnapshotId` - The snapshot ID to delete

**Emits:** Logs a `snapshot_restored` action to audit log

**Example:**
```typescript
traceabilityService.deleteSnapshot(snapshotId);
console.log('Snapshot deleted');
```

---

### compareSnapshots

Compares two snapshots and returns differences.

```typescript
compareSnapshots(id1: SnapshotId, id2: SnapshotId): Diff | null
```

**Parameters:**
- `id1: SnapshotId` - The first snapshot ID
- `id2: SnapshotId` - The second snapshot ID

**Returns:** `Diff | null` - The diff result, or null if either snapshot is not found

**Diff Structure:**
```typescript
interface Diff {
  added: string[];      // Reference IDs added in snapshot2
  removed: string[];    // Reference IDs removed from snapshot1
  modified: DiffChange[]; // Modified references
}

interface DiffChange {
  path: string;      // Reference ID
  oldValue: any;      // Value in snapshot1
  newValue: any;      // Value in snapshot2
}
```

**Example:**
```typescript
const diff = traceabilityService.compareSnapshots(snapshotId1, snapshotId2);
if (diff) {
  console.log(`Added: ${diff.added.length} references`);
  console.log(`Removed: ${diff.removed.length} references`);
  console.log(`Modified: ${diff.modified.length} references`);
}
```

---

### getReferenceHistory

Gets the change history for a reference.

```typescript
getReferenceHistory(refId: ReferenceId): DataReference['history']
```

**Parameters:**
- `refId: ReferenceId` - The reference ID

**Returns:** `ReferenceHistory[]` - Array of history entries

**Example:**
```typescript
const history = traceabilityService.getReferenceHistory(refId);
history.forEach((entry) => {
  console.log(`${entry.action}: ${entry.value} at ${new Date(entry.timestamp)}`);
});
```

---

### getCellHistory

Gets the change history for a cell.

```typescript
getCellHistory(cellId: CellId): CellChange[]
```

**Parameters:**
- `cellId: CellId` - The cell ID

**Returns:** `CellChange[]` - Array of cell change entries

**Example:**
```typescript
const history = traceabilityService.getCellHistory('sheet1!A1');
history.forEach((change) => {
  console.log(`${change.oldValue} -> ${change.newValue} by ${change.userId}`);
});
```

---

### updateReferenceHistory

Updates the history for a reference.

```typescript
updateReferenceHistory(refId: ReferenceId, history: DataReference['history']): void
```

**Parameters:**
- `refId: ReferenceId` - The reference ID
- `history: ReferenceHistory[]` - The new history array

**Note:** This is typically called internally by the `DataLinkManager`.

---

### recordCellChange

Records a cell change to history.

```typescript
recordCellChange(cellId: CellId, oldValue: any, newValue: any, userId: string): void
```

**Parameters:**
- `cellId: CellId` - The cell ID
- `oldValue: any` - The previous value
- `newValue: any` - The new value
- `userId: string` - The user who made the change

**Example:**
```typescript
traceabilityService.recordCellChange(
  'sheet1!A1',
  oldCellValue,
  newCellValue,
  'user-123'
);
```

---

### getDependents

Gets all references that depend on a specific cell.

```typescript
getDependents(cellId: CellId): ReferenceId[]
```

**Parameters:**
- `cellId: CellId` - The cell ID to check

**Returns:** `ReferenceId[]` - Array of reference IDs that depend on the cell

**Example:**
```typescript
const dependents = traceabilityService.getDependents('sheet1!A1');
console.log(`Cell has ${dependents.length} dependents`);
```

---

### getDependencies

Gets all data sources referenced by a document.

```typescript
getDependencies(docId: DocumentId): DataSource[]
```

**Parameters:**
- `docId: DocumentId` - The document ID

**Returns:** `DataSource[]` - Array of data sources the document references

**Example:**
```typescript
const dependencies = traceabilityService.getDependencies('doc-456');
console.log(`Document depends on ${dependencies.length} sources`);
dependencies.forEach((source) => {
  console.log(`- ${source.fileName} / ${source.sheetName}`);
});
```

---

### logAction

Logs an audit action.

```typescript
logAction(action: AuditAction): void
```

**Parameters:**
- `action: AuditAction` - The action to log

**AuditAction Structure:**
```typescript
interface AuditAction {
  action: AuditActionType;              // The action type
  userId?: string;                      // Optional user ID
  metadata?: Record<string, any>;        // Optional metadata
}
```

**AuditActionType Values:**
- `reference_created` - A reference was created
- `reference_updated` - A reference was updated
- `reference_deleted` - A reference was deleted
- `reference_refreshed` - A reference was refreshed
- `snapshot_created` - A snapshot was created
- `snapshot_restored` - A snapshot was deleted/restored
- `document_saved` - A document was saved
- `data_changed` - Data changed

**Example:**
```typescript
traceabilityService.logAction({
  action: 'reference_created',
  userId: 'user-123',
  metadata: { referenceId: 'ref-456' }
});
```

---

### getAuditLog

Gets audit logs with optional filters.

```typescript
getAuditLog(filters?: AuditFilters): AuditLog[]
```

**Parameters:**
- `filters: AuditFilters` (optional) - Filters to apply

**AuditFilters Structure:**
```typescript
interface AuditFilters {
  startDate?: number;      // Filter by start date (timestamp)
  endDate?: number;        // Filter by end date (timestamp)
  action?: AuditActionType; // Filter by action type
  userId?: string;         // Filter by user ID
  limit?: number;          // Limit number of results
}
```

**Returns:** `AuditLog[]` - Array of audit log entries, newest first

**Example:**
```typescript
// Get all logs
const allLogs = traceabilityService.getAuditLog();

// Get logs for a specific user
const userLogs = traceabilityService.getAuditLog({ userId: 'user-123' });

// Get logs for a specific action
const createdLogs = traceabilityService.getAuditLog({ action: 'reference_created' });

// Get logs in a time range
const dateLogs = traceabilityService.getAuditLog({
  startDate: startDate.getTime(),
  endDate: endDate.getTime()
});

// Get limited results
const recentLogs = traceabilityService.getAuditLog({ limit: 50 });
```

---

### clearAuditLog

Clears all audit logs.

```typescript
clearAuditLog(): void
```

**Example:**
```typescript
traceabilityService.clearAuditLog();
console.log('Audit log cleared');
```

## Types

### AuditLog

```typescript
interface AuditLog {
  id: string;                          // Unique log entry ID
  timestamp: number;                    // Unix timestamp
  action: AuditActionType;              // The action performed
  userId: string;                      // User who performed the action
  metadata: Record<string, any>;        // Additional metadata
}
```

### Diff

```typescript
interface Diff {
  added: string[];      // Reference IDs added
  removed: string[];    // Reference IDs removed
  modified: DiffChange[]; // Modified references
}
```

### DiffChange

```typescript
interface DiffChange {
  path: string;    // Reference ID
  oldValue: any;   // Value in first snapshot
  newValue: any;   // Value in second snapshot
}
```

### AuditFilters

```typescript
interface AuditFilters {
  startDate?: number;
  endDate?: number;
  action?: AuditActionType;
  userId?: string;
  limit?: number;
}
```

### CellChange

```typescript
interface CellChange {
  timestamp: number;   // When the change occurred
  oldValue: any;      // The previous value
  newValue: any;      // The new value
  userId: string;      // User who made the change
}
```

## DefaultTraceabilityService

The default implementation stores data in memory.

```typescript
class DefaultTraceabilityService implements TraceabilityService {
  constructor(
    private getDocumentContent: (docId: DocumentId) => any,
    private getDocumentReferences: (docId: DocumentId) => DataReference[]
  )
  destroy(): void
}
```

### Constructor

```typescript
constructor(
  getDocumentContent: (docId: DocumentId) => any,
  getDocumentReferences: (docId: DocumentId) => DataReference[]
)
```

**Parameters:**
- `getDocumentContent` - Function to get document content for snapshots
- `getDocumentReferences` - Function to get references for snapshots

### destroy

Cleans up resources.

```typescript
destroy(): void
```

## Factory Function

```typescript
function createTraceabilityService(
  getDocumentContent: (docId: DocumentId) => any,
  getDocumentReferences: (docId: DocumentId) => DataReference[]
): TraceabilityService
```

Creates a new `TraceabilityService` instance.

**Example:**
```typescript
const traceabilityService = createTraceabilityService(
  (docId) => fetchDocumentContent(docId),
  (docId) => fetchDocumentReferences(docId)
);
```

## Usage Example: Complete Workflow

```typescript
import { createTraceabilityService } from '@/core/trace/TraceabilityService';

// Create service
const traceabilityService = createTraceabilityService(
  (docId) => getDocumentContent(docId),
  (docId) => getDocumentReferences(docId)
);

// Create a snapshot
const snapshotId1 = await traceabilityService.createSnapshot('doc-456');
console.log('Initial snapshot:', snapshotId1);

// Make some changes to the document...

// Create another snapshot
const snapshotId2 = await traceabilityService.createSnapshot('doc-456');
console.log('Second snapshot:', snapshotId2);

// Compare snapshots
const diff = traceabilityService.compareSnapshots(snapshotId1, snapshotId2);
if (diff) {
  console.log(`Changes detected:`);
  console.log(`  Added: ${diff.added.length} references`);
  console.log(`  Removed: ${diff.removed.length} references`);
  console.log(`  Modified: ${diff.modified.length} references`);
}

// Get reference history
const refId = 'ref-123';
const history = traceabilityService.getReferenceHistory(refId);
console.log(`Reference has ${history.length} history entries`);

// Get audit log
const logs = traceabilityService.getAuditLog({ limit: 10 });
console.log('Recent audit activity:');
logs.forEach((log) => {
  console.log(`  ${log.action} at ${new Date(log.timestamp).toLocaleString()}`);
});

// Get dependencies
const dependencies = traceabilityService.getDependencies('doc-456');
console.log(`Document depends on ${dependencies.length} data sources`);

// Get dependents
const cellId = 'sheet1!A1';
const dependents = traceabilityService.getDependents(cellId);
console.log(`Cell has ${dependents.length} dependent references`);

// Cleanup when done
traceabilityService.destroy();
```

## Automatic Event Logging

The `TraceabilityService` automatically logs events from the `EventBus`:

| Event | Logged Action |
|-------|---------------|
| `Events.ReferenceCreated` | `reference_created` |
| `Events.ReferenceUpdated` | `reference_updated` |
| `Events.ReferenceBroken` | `reference_updated` (with broken state) |
| `Events.DataChanged` | `data_changed` |
| `Events.DocumentSaved` | `document_saved` |

No manual event subscription is required for these automatic logs.

## Memory Management

The default implementation keeps:
- Snapshots: In-memory until destroyed
- Reference history: In-memory until destroyed
- Cell history: In-memory until destroyed
- Audit logs: Last 1000 entries (automatic cleanup)

For production use, consider implementing a persistent storage adapter.
