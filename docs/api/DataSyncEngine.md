# DataSyncEngine API

The `DataSyncEngine` is responsible for handling manual and automatic synchronization of data references between Word documents and Excel data sources.

## Overview

The `DataSyncEngine` provides functionality to:
- Refresh individual references or entire documents
- Enable/disable automatic synchronization
- Detect and resolve conflicts
- Notify listeners when data changes

## Installation

```typescript
import { DataSyncEngine, createDataSyncEngine } from '@/core/sync/DataSyncEngine';

// Create a new instance
const syncEngine = createDataSyncEngine(dataLinkManager);
```

## Interface

```typescript
interface DataSyncEngine {
  // Manual Refresh
  refreshReference(id: ReferenceId): Promise<void>;
  refreshDocument(docId: DocumentId): Promise<void>;
  refreshAll(): Promise<void>;
  refreshBySource(source: DataSource): Promise<void>;

  // Auto Sync
  enableAutoSync(enabled: boolean): void;
  setSyncInterval(interval: number): void;
  isAutoSyncEnabled(): boolean;

  // Conflict Detection
  detectConflicts(): SyncConflict[];
  resolveConflict(conflict: SyncConflict, strategy: ResolveStrategy): void;

  // Event Handlers
  onDataChange(callback: (changes: DataChange[]) => void): void;
  offDataChange(callback: (changes: DataChange[]) => void): void;

  // Cleanup
  destroy(): void;
}
```

## Methods

### refreshReference

Refreshes a single reference to get the latest value from the data source.

```typescript
refreshReference(id: ReferenceId): Promise<void>
```

**Parameters:**
- `id: ReferenceId` - The reference ID to refresh

**Returns:** `Promise<void>`

**Emits:** `Events.DataChanged` (if value changes)

**Example:**
```typescript
await syncEngine.refreshReference(refId);
console.log('Reference refreshed');
```

---

### refreshDocument

Refreshes all references in a document.

```typescript
refreshDocument(docId: DocumentId): Promise<void>
```

**Parameters:**
- `docId: DocumentId` - The document ID to refresh

**Returns:** `Promise<void>`

**Emits:** `Events.DataChanged` (if any values change)

**Example:**
```typescript
await syncEngine.refreshDocument('doc-456');
console.log('Document refreshed');
```

---

### refreshAll

Refreshes all references across all documents.

```typescript
refreshAll(): Promise<void>
```

**Returns:** `Promise<void>`

**Emits:** `Events.DataChanged` (if any values change)

**Example:**
```typescript
await syncEngine.refreshAll();
console.log('All references refreshed');
```

---

### refreshBySource

Refreshes all references that use the given data source (selective refresh by source).

```typescript
refreshBySource(source: DataSource): Promise<void>
```

**Parameters:**
- `source: DataSource` - The data source (fileId, sheetId, range, etc.); only references matching this source are refreshed

**Returns:** `Promise<void>`

**Emits:** `Events.DataChanged` (if any values change)

**Example:**
```typescript
await syncEngine.refreshBySource({
  fileId: 'file-123',
  fileName: 'budget.xlsx',
  sheetId: 'sheet-1',
  sheetName: 'Sheet1',
  range: { startRow: 0, startCol: 0 },
  isFormula: false,
});
```

---

### enableAutoSync

Enables or disables automatic synchronization.

```typescript
enableAutoSync(enabled: boolean): void
```

**Parameters:**
- `enabled: boolean` - Whether to enable auto-sync

**Example:**
```typescript
// Enable auto-sync
syncEngine.enableAutoSync(true);

// Disable auto-sync
syncEngine.enableAutoSync(false);
```

**Note:** When auto-sync is enabled, the engine will periodically call `refreshAll()` at the configured interval.

---

### setSyncInterval

Sets the auto-sync interval in milliseconds.

```typescript
setSyncInterval(interval: number): void
```

**Parameters:**
- `interval: number` - The sync interval in milliseconds (default: 60000 = 1 minute)

**Example:**
```typescript
// Sync every 30 seconds
syncEngine.setSyncInterval(30000);

// Sync every 5 minutes
syncEngine.setSyncInterval(300000);
```

---

### isAutoSyncEnabled

Checks if auto-sync is currently enabled.

```typescript
isAutoSyncEnabled(): boolean
```

**Returns:** `boolean` - Whether auto-sync is enabled

**Example:**
```typescript
if (syncEngine.isAutoSyncEnabled()) {
  console.log('Auto-sync is active');
}
```

---

### detectConflicts

Detects conflicts in references.

```typescript
detectConflicts(): SyncConflict[]
```

**Returns:** `SyncConflict[]` - Array of detected conflicts

**Example:**
```typescript
const conflicts = syncEngine.detectConflicts();
if (conflicts.length > 0) {
  console.warn(`Found ${conflicts.length} conflicts`);
  conflicts.forEach(conflict => {
    console.log('Conflict:', conflict.referenceId);
  });
}
```

---

### resolveConflict

Resolves a conflict with the specified strategy.

```typescript
resolveConflict(conflict: SyncConflict, strategy: ResolveStrategy): void
```

**Parameters:**
- `conflict: SyncConflict` - The conflict to resolve
- `strategy: ResolveStrategy` - The resolution strategy

**Returns:** `void`

**Emits:** `Events.ReferenceUpdated`

**Strategies:**
- `'keep-local'`: Keep the local (current) value
- `'keep-remote'`: Use the remote (source) value
- `'merge'`: Merge values (timestamp-based)

**Example:**
```typescript
syncEngine.resolveConflict(conflict, 'keep-remote');
console.log('Conflict resolved');
```

---

### onDataChange

Registers a callback for data change notifications.

```typescript
onDataChange(callback: (changes: DataChange[]) => void): void
```

**Parameters:**
- `callback: (changes: DataChange[]) => void` - The callback function

**Example:**
```typescript
syncEngine.onDataChange((changes) => {
  console.log('Data changed:', changes);
  changes.forEach(change => {
    console.log(`Reference ${change.referenceId}: ${change.oldValue} -> ${change.newValue}`);
  });
});
```

---

### offDataChange

Unregisters a data change callback (pair of onDataChange).

```typescript
offDataChange(callback: (changes: DataChange[]) => void): void
```

**Parameters:**
- `callback: (changes: DataChange[]) => void` - The callback to remove

**Example:**
```typescript
const handler = (changes) => console.log(changes);
syncEngine.onDataChange(handler);
syncEngine.offDataChange(handler);
```

---

### destroy

Cleans up resources and stops auto-sync.

```typescript
destroy(): void
```

**Example:**
```typescript
syncEngine.destroy();
console.log('Sync engine destroyed');
```

---

## DefaultDataSyncEngine

The default implementation of `DataSyncEngine`.

```typescript
class DefaultDataSyncEngine implements DataSyncEngine {
  constructor(
    private dataLinkManager: DataLinkManager,
    private dataSourceResolver: (ref: DataReference) => Promise<any>
  )
}
```

### Constructor

```typescript
constructor(
  dataLinkManager: DataLinkManager,
  dataSourceResolver?: (ref: DataReference) => Promise<any>
)
```

**Parameters:**
- `dataLinkManager: DataLinkManager` - The data link manager instance
- `dataSourceResolver` (optional) - Custom function to resolve data sources

## Factory Function

```typescript
function createDataSyncEngine(
  dataLinkManager: DataLinkManager,
  dataSourceResolver?: (ref: DataReference) => Promise<any>
): DataSyncEngine
```

Creates a new `DataSyncEngine` instance.

**Example:**
```typescript
const syncEngine = createDataSyncEngine(dataLinkManager);
```

## Events

The `DataSyncEngine` uses the following events:

| Event | When Emitted | Description |
|-------|---------------|-------------|
| `Events.DataChanged` | Reference value changes | Emits array of data changes when references are refreshed |
| `Events.ReferenceUpdated` | Conflict resolution | Emits when a conflict is resolved |

## Conflict Types

```typescript
type ResolveStrategy = 'keep-local' | 'keep-remote' | 'merge';

interface Conflict {
  id: string;
  type: ConflictType;
  referenceId: string;
  reference: DataReference;
  localValue: any;
  remoteValue: any;
  localTimestamp: number;
  remoteTimestamp: number;
  resolved: boolean;
}

type ConflictType = 'data-inconsistency' | 'version-mismatch' | 'concurrent-edit';
```

## Usage Example: Complete Workflow

```typescript
import { createDataLinkManager } from '@/core/editors/DataLinkManager';
import { createDataSyncEngine } from '@/core/sync/DataSyncEngine';
import { eventBus, Events } from '@/core/eventBus';

// Create managers
const dataLinkManager = createDataLinkManager();
const syncEngine = createDataSyncEngine(dataLinkManager);

// Listen for data changes via event bus
eventBus.on(Events.DataChanged, (changes) => {
  console.log('Data changed via event bus:', changes);
});

// Listen for data changes via sync engine
syncEngine.onDataChange((changes) => {
  console.log('Data changed via sync engine:', changes);
});

// Enable auto-sync
syncEngine.setSyncInterval(30000); // 30 seconds
syncEngine.enableAutoSync(true);

// Refresh a single reference
await syncEngine.refreshReference(refId);

// Refresh a document
await syncEngine.refreshDocument('doc-456');

// Detect conflicts
const conflicts = syncEngine.detectConflicts();
if (conflicts.length > 0) {
  // Resolve each conflict
  conflicts.forEach(conflict => {
    syncEngine.resolveConflict(conflict, 'keep-remote');
  });
}

// Disable auto-sync
syncEngine.enableAutoSync(false);

// Cleanup
syncEngine.destroy();
```

## Data Change Structure

```typescript
interface DataChange {
  referenceId: string;
  oldValue: any;
  newValue: any;
  timestamp: number;
  source: 'manual' | 'auto';
}
```

- `referenceId`: The ID of the changed reference
- `oldValue`: The previous value
- `newValue`: The new value
- `timestamp`: When the change occurred (Unix timestamp)
- `source`: Whether the change was triggered manually or automatically
