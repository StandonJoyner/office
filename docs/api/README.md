# Office Suite API Documentation

This directory contains API documentation for the Office Suite application's core modules.

## Overview

The Office Suite is built with a modular architecture that separates concerns across several key areas:

- **DataLinkManager**: Manages references between Excel data sources and Word documents (create/resolve/query; use `createDataLinkManager()`).
- **DataSyncEngine**: Handles manual and automatic synchronization of data references (refresh by reference/document/source/all; use `createDataSyncEngine(dataLinkManager[, resolver])`).
- **TraceabilityService**: Provides version tracking, change history, dependency analysis, and audit logging (use `createTraceabilityService(docContentLoader, docRefsLoader)`).
- **EventBus**: Singleton event bus for publish-subscribe communication; use `eventBus` and `Events` enum (`reference:created`, `data:changed`, etc.).
- **Core Types**: Shared type definitions (`DataReference`, `DataSource`, `ReferenceTarget`, etc.) from `@/core/types`.

## Table of Contents

- [DataLinkManager API](./DataLinkManager.md)
- [DataSyncEngine API](./DataSyncEngine.md)
- [TraceabilityService API](./TraceabilityService.md)
- [EventBus](./EventBus.md)
- [Type Definitions](./Types.md)

## Quick Start

```typescript
import {
  createDataLinkManager,
  createDataSyncEngine,
  createTraceabilityService,
  eventBus,
  Events
} from '@/core';

// Create managers
const dataLinkManager = createDataLinkManager();
const syncEngine = createDataSyncEngine(dataLinkManager);
const traceabilityService = createTraceabilityService(
  () => getDocumentContent(docId),
  () => getDocumentReferences(docId)
);

// Listen for events
eventBus.on(Events.DataChanged, (changes) => {
  console.log('Data changed:', changes);
});

// Create a reference
const source = {
  fileId: 'file-123',
  fileName: 'budget.xlsx',
  sheetId: 'sheet-1',
  sheetName: 'Sheet1',
  range: { startRow: 0, startCol: 0 },
  isFormula: false
};

const refId = dataLinkManager.createReference(source, 'doc-456');
```

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                         UI Layer                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐ │
│  │ WordEditor   │  │ ExcelEditor  │  │  Panels      │ │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘ │
│         │                    │                    │           │
└─────────┼────────────────────┼────────────────────┼───────────┘
          │                    │                    │
┌─────────┴────────────────────┴────────────────────┴───────────┐
│                    Core Layer                             │
│  ┌──────────────────┐  ┌──────────────────┐  ┌─────────┐ │
│  │ DataLinkManager │  │  DataSyncEngine │  │ EventBus │ │
│  └────────┬─────────┘  └────────┬─────────┘  └────┬────┘ │
│           │                      │                    │     │
│  ┌────────┴──────────────────────┴────────────────────┴────┐ │
│  │          TraceabilityService                          │ │
│  └────────────────────────────────────────────────────────────┘ │
│                                                              │
│  ┌──────────────────────────────────────────────────────────────┐ │
│  │                    Data Layer                           │ │
│  │  (File System, Storage, System)                     │ │
│  └──────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────┘
```

## Event Flow

```
Excel Cell Changed
       │
       ▼
ExcelDataSource detects change
       │
       ▼
DataLinkManager finds affected references
       │
       ▼
DataLinkManager resolves and updates references
       │
       ▼
EventBus emits DataChanged event
       │
       ├────────────────────────────────┐
       ▼                            ▼
DataSyncEngine              UI Components
subscribes and               (ReferenceNodeView,
processes changes              TraceabilityPanel)
       │
       ▼
EventBus emits ReferenceUpdated event
       │
       ▼
TraceabilityService logs to audit log
```

## Conventions

### ID Generation

All IDs are generated using `crypto.randomUUID()`:

```typescript
const id = crypto.randomUUID();
```

### Event Naming

Events use a hierarchical naming convention: `category:action`. Use the `Events` enum when subscribing.

```typescript
Events.ReferenceCreated  // 'reference:created' - A new reference was created
Events.ReferenceUpdated  // 'reference:updated' - A reference was updated
Events.DataChanged       // 'data:changed'     - Data values changed
Events.DocumentSaved     // 'document:saved'    - Document was saved
```

### State Management

References have four possible states:

- `active`: Reference is up-to-date and valid
- `stale`: Reference may need updating (source changed recently)
- `broken`: Reference source is no longer available
- `conflict`: Reference has conflicting data
