# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
claude --resume cc0c4776-5171-4096-9301-36159377ccc1

## Common Commands

### Development
```bash
npm run dev          # Start Next.js dev server (defaults to port 3000)
npm run build        # Build for production (includes type checking)
npm run start        # Start production server
npm run lint         # Run ESLint
npm run type-check   # Run standalone TypeScript type checking
npm run format       # Format code with Prettier
```

**Note**: `npm run type-check` is a standalone check. The `npm run build` command includes type checking via Next.js. The `tsc --noEmit` approach is not configured - use `npm run build` for type errors.

### Testing
```bash
npm test            # Run tests with Vitest
npm run test:ui     # Run tests with Vitest UI
```

### Desktop (Tauri)
```bash
npm run tauri:dev        # Start Tauri dev environment
npm run tauri:build       # Build Tauri desktop app
npm run tauri:build-debug # Build Tauri debug version
```

### Running Single Test
Use Vitest's `-t` flag or select tests in the UI:
```bash
npm test -- -t "test name"
```

## Architecture Overview

This is an office suite with deep integration between Word and Excel editors, featuring data traceability and dynamic updates. The architecture follows a **web/desktop same-source** pattern using an adapter layer.

### Core Layers

**Presentation Layer** (`ui/`, `app/`, `components/`)
- React components for Word and Excel editors
- Shared UI components (TraceabilityPanel, DataUpdateNotification)
- React hooks (`ui/hooks/`) for integrating with core services

**Editor Core Layer**
- Word: Tiptap (ProseMirror-based) in `ui/word/`
  - `ReferenceNode`: Custom atom node for single cell references
  - `RangeTableNode`: Custom table node for Excel range references (extends @tiptap/extension-table)
- Excel: Univer in `ui/excel/` (ExcelEditor)

**Business Logic Layer** (`core/`)
- `eventBus.ts`: Event-driven communication between modules
- `types.ts`: Core type definitions (`DataReference`, `DataSource`, `ReferenceState`, `RangeTableMeta`, etc.)
- `editors/DataLinkManager.ts`: Manages references between Excel data sources and Word documents
- `editors/ExcelDataSource.ts`: Excel cell/range resolution and change monitoring
- `editors/UniverWorkbookAdapter.ts`: Adapter for Univer workbook API
- `sync/DataSyncEngine.ts`: Data synchronization with conflict resolution (exported from core/index.ts)
- `sync/DataBinding.ts`: Reactive two-way binding using Signal pattern
- `sync/ConflictResolver.ts`: Conflict detection and resolution strategies
- `sync/UpdateQueue.ts`: Batching, deduplication, retry logic for updates
- `sync/PerformanceMonitor.ts`: Performance metrics and monitoring
- `trace/TraceabilityService.ts`: Version tracking, change history, dependency analysis

**Data Layer** (`adapters/`)
- Platform-agnostic interfaces for file system, storage, and system operations
- Web implementations in `adapters/web/`
- Tauri desktop implementations in `adapters/tauri/`
- Runtime detection and adapter factory in `adapters/index.ts`

### Event-Driven Architecture

The system uses an event bus (`core/eventBus.ts`) for cross-module communication. Events are defined in the `Events` enum:

```typescript
enum Events {
  ReferenceCreated = 'reference:created',
  ReferenceUpdated = 'reference:updated',
  ReferenceDeleted = 'reference:deleted',
  ReferenceBroken = 'reference:broken',
  DataChanged = 'data:changed',
  DocumentSaved = 'document:saved',
}
```

Key pattern: When Excel data changes, `ExcelDataSource` emits `Events.DataChanged`, which `DataLinkManager` listens to and triggers updates to affected references. Range tables in auto-sync mode also listen to `DataChanged` events via `useEditorDataSync` to update table content automatically.

### Data Reference Flow

#### Single Cell References
1. **Creation**: User selects a cell in Excel → `DataLinkManager.createReference()` creates a reference
2. **Storage**: Reference stored in `MemoryDataLinkManager` with source (Excel) and target (Word) info
3. **Resolution**: `ExcelDataSourceManager.resolveReference()` fetches current value from Univer
4. **Sync**: `DataSyncEngine` handles manual/automatic refresh with conflict resolution
5. **Traceability**: `TraceabilityService` tracks versions, history, and dependencies

#### Range Reference Tables
1. **Creation**: User selects a range in Excel → `WordEditor.insertRangeTable()` creates a table reference
2. **Storage**: Reference stored with `tableMeta` (rowCount, colCount, syncMode, preserveFormatting)
3. **Insertion**: Table rows are built from 2D array data and inserted as `rangeTable` node
4. **Sync**: `useEditorDataSync` handles table content updates on data changes
5. **State Management**: Table state (active/stale/broken/conflict) displayed and updated via event bus

### Reactive Signal Pattern (DataBinding)

`core/sync/DataBinding.ts` implements a reactive signal pattern similar to Preact Signals or Angular Signals:

- **Signal<T>**: Reactive value holder that notifies subscribers when value changes
- **ComputedSignal<T>**: Derived value that automatically recomputes when dependencies change
- **DataBindingManager**: Manages bindings using signals for reactive updates

Key pattern:
```typescript
// Create a signal
const signal = new Signal(initialValue);

// Subscribe to changes
const unsubscribe = signal.subscribe(newValue => console.log(newValue));

// Update value (notifies all subscribers)
signal.set(newValue);
```

**Important**: `DataBindingManager` is **NOT exported** from `core/index.ts`. Import directly:
```typescript
import { DataBindingManager, createDataBindingManager } from '@/core/sync/DataBinding';
```

### Module Export Clarification

Not all `core/sync/` modules are re-exported from `core/index.ts`. Check import locations:

| Module | Exported from core/index.ts? | Import from |
|--------|----------------------------|-------------|
| DataSyncEngine | Yes | `@/core` or `@/core/sync/DataSyncEngine` |
| ConflictResolver | No | `@/core/sync/ConflictResolver` |
| UpdateQueue | No | `@/core/sync/UpdateQueue` |
| DataBindingManager | No | `@/core/sync/DataBinding` |
| PerformanceMonitor | No | `@/core/sync/PerformanceMonitor` |

**DataLinkManager vs DataBindingManager**:
- `DataLinkManager`: Manages references, creates/updates/deletes them, resolves values from Excel
- `DataBindingManager`: Provides reactive two-way binding using signals, subscribes to value changes
- They serve different purposes - don't confuse them!

### Global Batch Update State

`ui/hooks/useBatchUpdate.ts` exports global utilities for queuing updates:

```typescript
import { enqueueUpdate, getPendingUpdateCount, clearPendingUpdates, getAllPendingUpdates } from '@/ui/hooks/useBatchUpdate';

// Queue an update for later processing
enqueueUpdate('reference-id', { some: 'data' });

// Get pending count
const count = getPendingUpdateCount();
```

This pattern allows batching updates across component boundaries.

### Platform Adapters

The `adapters/` directory implements adapter pattern for Web/Desktop differences:

| Interface | Web Implementation | Desktop (Tauri) |
|------------|-------------------|---------------------|
| `FileAdapter` | `web/FileAdapterWeb.ts` | `tauri/FileAdapterTauri.ts` |
| `SystemAdapter` | `web/SystemAdapterWeb.ts` | `tauri/SystemAdapterTauri.ts` |
| `StorageAdapter` | (uses IndexedDB) | `tauri/storage.ts` |

The `adapters/index.ts` factory detects runtime environment and provides appropriate adapter.

### Key Types (`core/types.ts`)

- `DataReference`: Core reference linking Excel source to Word target
- `DataSource`: Excel file/sheet/cell reference info
- `CellRange`: Excel range coordinates (startRow, startCol, endRow?, endCol?)
- `ReferenceTarget`: Document/node position in Word
- `ReferenceState`: `'active' | 'stale' | 'broken' | 'conflict'`
- `RangeTableMeta`: Table-specific metadata including rowCount, colCount, syncMode, preserveFormatting
- `ReferenceDisplay.tableMeta`: Optional table metadata for range references

### Univer Excel Integration Notes

- Uses preset-based initialization: `createUniver({ presets: [UniverSheetsCorePreset()] })`
- Listen to cell changes via `'sheet.operation'` events
- Get/set cells using `FUniver` facade API: `api.getActiveWorkbook().getActiveSheet().getRange()`
- Export `ExcelEditorRefType` interface for parent component access to methods like `getCellValue()`, `setCellValue()`

### Tiptap Word Integration Notes

- `ReferenceNode`: Custom Tiptap atom node for inline data references
- `ReferenceNodeView.tsx`: React component rendering the reference with state visualization
- `RangeTableNode`: Custom Tiptap table node for Excel range references (extends @tiptap/extension-table)
- `RangeTableNodeView.tsx`: React component rendering the range table with state visualization, sync controls, and delete functionality
- References use inline atom nodes with HTML serialization/deserialization
- Range references insert as tables with fixed size determined at creation time

### Range Reference Table Features

- **Table Insertion**: Excel cell ranges insert as Tiptap tables in Word documents
- **Sync Modes**: Manual (default) or Auto-sync for automatic updates
- **State Indicators**: Visual indicators for active, stale, broken, and conflict states
- **Interactive Controls**:
  - Refresh button for manual sync
  - Sync mode toggle (manual ↔ auto)
  - Delete button with confirmation
- **Hover Toolbar**: Shows controls when hovering over table
- **Selection Ring**: Visual indicator when table is selected
- **Source Tracking**: Stores Excel file, sheet, and range information in table attributes

### WordEditor Ref API

```typescript
// Reference methods
editorRef.current?.insertReference(reference: DataReference);
editorRef.current?.insertRangeTable(reference: DataReference);
editorRef.current?.getEditor(); // Returns Tiptap editor instance
editorRef.current?.getReferenceIds(); // Returns array of all reference IDs
```

**Key methods**:
- `insertReference()`: Insert a single cell reference at cursor position
- `insertRangeTable()`: Insert a range reference as table at cursor position
- `getEditor()`: Access underlying Tiptap editor for advanced operations
- `getReferenceIds()`: Get all reference IDs for document-wide operations

### Test Structure

Test files are located alongside their source files:
- `core/eventBus.test.ts`
- `core/editors/DataLinkManager.test.ts`
- `core/editors/ExcelDataSource.test.ts` (includes range resolution tests)
- `core/sync/DataSyncEngine.test.ts`
- `core/trace/TraceabilityService.test.ts`
- `ui/word/RangeTableNode.test.ts` (node type tests)
- `ui/word/RangeTableNodeView.test.tsx` (E2E component tests - 24 tests)
- `ui/hooks/useEditorDataSync.test.ts` (sync hook tests including range tables)

Tests use Vitest with jsdom environment configured in `vitest.config.ts` and setup file `vitest.setup.ts` for jest-dom matchers.

## Important Notes

- Not all modules export from `core/index.ts` - check `Module Export Clarification` section above
- The dev server may start on port 3001 if 3000 is in use
- Tests use Vitest with jsdom environment
- TypeScript strict mode is enabled with `downlevelIteration: true` for `for...of` loops on Maps
- Univer API event types may require `as any` type assertions due to API limitations
- ExcelEditor uses `React.forwardRef` for ref access - import as `ExcelEditor` with ref prop
- Use `crypto.randomUUID()` for generating IDs (no external dependency)
- `reference.metadata` can be undefined - always use optional chaining: `reference.metadata?.updated`
- For type checking, prefer `npm run build` over `npm run type-check` (the standalone tsc check may not work properly)

## Common Patterns and Pitfalls

### ID Generation
Always use `crypto.randomUUID()` for generating IDs:
```typescript
const id = crypto.randomUUID();
```

### Optional Metadata
Reference metadata is optional, use nullish coalescing:
```typescript
const updated = ref.metadata?.updated ?? Date.now();
```

### Event Bus Usage
```typescript
import { eventBus, Events } from '@/core/eventBus';

// Listen to events
const unsubscribe = eventBus.on(Events.DataChanged, (data) => {
  console.log(data);
});

// Emit events
eventBus.emit(Events.ReferenceCreated, reference);
```

### Factory Functions
Most core modules export factory functions:
```typescript
import { createDataLinkManager, createDataSyncEngine, createTraceabilityService } from '@/core';

const manager = createDataLinkManager();
const syncEngine = createDataSyncEngine(manager);
const traceService = createTraceabilityService(docLoader, refLoader);
```

## Range Reference Implementation

### Overview
Range references allow Excel cell ranges to be inserted as tables in Word documents with automatic data synchronization.

### Key Components

**RangeTableNode** (`ui/word/RangeTableNode.ts`)
- Extends @tiptap/extension-table to preserve native Tiptap table functionality
- Custom attributes: refId, syncMode, sourceInfo
- Supports HTML serialization/deserialization for document persistence

**RangeTableNodeView** (`ui/word/RangeTableNodeView.tsx`)
- React component rendering the range table with UI controls
- State indicators (active, stale, broken, conflict) with visual icons
- Hover toolbar with refresh, sync toggle, and delete buttons
- Auto-sync mode subscribes to DataChanged events

**useEditorDataSync** (`ui/hooks/useEditorDataSync.ts`)
- Handles both single cell references and range table references
- On DataChanged events, updates matching range tables with new cell data
- On ReferenceUpdated events, refreshes table content and updates state

**WordEditor** (`ui/word/WordEditor.tsx`)
- `insertRangeTable()`: Insert a range reference as table at cursor position
- `onInsertTableReferenceRequest`: Callback prop for parent to handle table reference creation
- "+ Table Reference" button in toolbar to trigger table reference insertion

### Testing

**Unit Tests**: 78 tests passing
- RangeTableNode: 2 tests (node type validation)
- RangeTableNodeView: 24 E2E tests (all user interactions)
- ExcelDataSource: 43 tests (including 6 range resolution tests)
- DataLinkManager: 39 tests (including 3 range reference tests)
- useEditorDataSync: 12 tests (sync functionality)

**Manual Testing**: See `docs/plans/manual-testing-checklist.md` for comprehensive manual testing guide with 80+ test cases.

### Documentation

- `docs/plans/implementation-summary.md`: Complete implementation overview and architecture diagram
- `docs/plans/manual-testing-checklist.md`: Manual testing checklist for QA verification
