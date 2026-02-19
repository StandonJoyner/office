# EventBus

The `EventBus` provides a publish-subscribe pattern for event-driven communication between modules.

## Overview

The `EventBus` allows components to:
- Subscribe to events
- Unsubscribe from events
- Emit events with data
- Clear all event listeners

## Installation

```typescript
import { eventBus, Events } from '@/core/eventBus';

// Use the singleton instance
eventBus.on(Events.DataChanged, (data) => {
  console.log('Data changed:', data);
});
```

## Events Enum

```typescript
enum Events {
  // Reference Events
  ReferenceCreated = 'reference:created',
  ReferenceUpdated = 'reference:updated',
  ReferenceDeleted = 'reference:deleted',
  ReferenceBroken = 'reference:broken',

  // Data Events
  DataChanged = 'data:changed',

  // Document Events
  DocumentSaved = 'document:saved',

  // Excel Events
  ExcelWorkbookLoaded = 'excel:workbook:loaded',
  ExcelCellSelected = 'excel:cell:selected',
  ExcelRangeSelected = 'excel:range:selected',
}
```

## Methods

### on

Subscribes a handler to an event.

```typescript
on<T = any>(event: Events, handler: EventHandler<T>): void
```

**Parameters:**
- `event: Events` - The event to subscribe to
- `handler: EventHandler<T>` - The callback function

**Example:**
```typescript
// Subscribe to data changes
eventBus.on(Events.DataChanged, (changes: DataChange[]) => {
  console.log('Data changed:', changes);
});

// Subscribe to reference creation
eventBus.on(Events.ReferenceCreated, (ref: DataReference) => {
  console.log('New reference created:', ref.id);
});

// Subscribe with custom data type
eventBus.on(Events.ExcelCellSelected, (data: { sheetId: string, row: number, col: number }) => {
  console.log(`Cell selected: ${data.sheetId}!${data.row}:${data.col}`);
});
```

---

### off

Unsubscribes a handler from an event.

```typescript
off<T = any>(event: Events, handler: EventHandler<T>): void
```

**Parameters:**
- `event: Events` - The event to unsubscribe from
- `handler: EventHandler<T>` - The callback function to remove

**Example:**
```typescript
const handler = (changes: DataChange[]) => {
  console.log('Data changed:', changes);
};

// Subscribe
eventBus.on(Events.DataChanged, handler);

// Later, unsubscribe
eventBus.off(Events.DataChanged, handler);
```

---

### emit

Emits an event with data to all subscribed handlers.

```typescript
emit<T = any>(event: Events, payload: T): void
```

**Parameters:**
- `event: Events` - The event to emit
- `payload: T` - The data to pass to handlers

**Example:**
```typescript
// Emit data changes
eventBus.emit(Events.DataChanged, [
  {
    referenceId: 'ref-123',
    oldValue: 100,
    newValue: 200,
    timestamp: Date.now(),
    source: 'manual'
  }
]);

// Emit reference update
eventBus.emit(Events.ReferenceUpdated, {
  referenceId: 'ref-123',
  action: 'refresh',
  oldState: { /* ... */ },
  newState: { /* ... */ }
});
```

---

### clear

Clears all event listeners.

```typescript
clear(): void
```

**Example:**
```typescript
eventBus.clear();
console.log('All event listeners cleared');
```

## Event Handler Type

```typescript
type EventHandler<T = any> = (data: T) => void;
```

## Event Payloads

### ReferenceCreated

```typescript
{
  reference: DataReference;
}
```

### ReferenceUpdated

```typescript
{
  referenceId: string;
  oldState?: DataReference;
  newState: DataReference;
  action?: 'refresh' | string;
}
```

### ReferenceDeleted

```typescript
{
  referenceId: string;
}
```

### ReferenceBroken

```typescript
{
  referenceId: string;
  state: 'broken';
}
```

### DataChanged

```typescript
{
  referenceId: string;
  oldValue: any;
  newValue: any;
  timestamp: number;
  source: 'manual' | 'auto';
}[]
```

### DocumentSaved

```typescript
{
  documentId: string;
  timestamp?: number;
}
```

### ExcelWorkbookLoaded

```typescript
{
  workbookId: string;
  sheetCount: number;
}
```

### ExcelCellSelected

```typescript
{
  sheetId: string;
  row: number;
  col: number;
  value?: any;
}
```

### ExcelRangeSelected

```typescript
{
  sheetId: string;
  range: CellRange;
  values?: any[][];
}
```

## Usage Example: Complete Workflow

```typescript
import { eventBus, Events } from '@/core/eventBus';

// Define handlers
const dataChangeHandler = (changes: DataChange[]) => {
  console.log('Data changed:', changes.length, 'references');
};

const referenceCreatedHandler = (data: { reference: DataReference }) => {
  console.log('Reference created:', data.reference.id);
};

const cellSelectedHandler = (data: { sheetId: string, row: number, col: number }) => {
  console.log(`Cell selected: ${data.sheetId}!${data.row}:${data.col}`);
};

// Subscribe to events
eventBus.on(Events.DataChanged, dataChangeHandler);
eventBus.on(Events.ReferenceCreated, referenceCreatedHandler);
eventBus.on(Events.ExcelCellSelected, cellSelectedHandler);

// Emit events
eventBus.emit(Events.DataChanged, [
  {
    referenceId: 'ref-123',
    oldValue: 100,
    newValue: 200,
    timestamp: Date.now(),
    source: 'manual'
  }
]);

// Unsubscribe when component unmounts
eventBus.off(Events.DataChanged, dataChangeHandler);
eventBus.off(Events.ReferenceCreated, referenceCreatedHandler);
eventBus.off(Events.ExcelCellSelected, cellSelectedHandler);
```

## React Integration Example

```typescript
'use client';

import { useEffect } from 'react';
import { eventBus, Events } from '@/core/eventBus';

function MyComponent() {
  useEffect(() => {
    // Subscribe to data changes
    const handler = (changes: DataChange[]) => {
      console.log('Data changed:', changes);
    };

    eventBus.on(Events.DataChanged, handler);

    // Cleanup on unmount
    return () => {
      eventBus.off(Events.DataChanged, handler);
    };
  }, []);

  // Emit an event
  const handleUpdate = () => {
    eventBus.emit(Events.ReferenceUpdated, {
      referenceId: 'ref-123',
      action: 'refresh'
    });
  };

  return <button onClick={handleUpdate}>Update Reference</button>;
}
```

## Best Practices

### 1. Always Unsubscribe

Always clean up event listeners to prevent memory leaks:

```typescript
useEffect(() => {
  const handler = (data) => { /* ... */ };
  eventBus.on(Events.DataChanged, handler);

  return () => {
    eventBus.off(Events.DataChanged, handler);
  };
}, []);
```

### 2. Error Handling

The `EventBus` catches errors in handlers and logs them:

```typescript
// Errors in handlers won't crash the app
eventBus.on(Events.DataChanged, (data) => {
  throw new Error('This will be caught and logged');
});
// Console output: "Error in event handler for data:changed: Error: This will be caught and logged"
```

### 3. Event Naming

Use the predefined `Events` enum for type safety:

```typescript
// Good - type-safe
eventBus.on(Events.DataChanged, handler);

// Bad - no type safety
eventBus.on('data:changed', handler);
```

### 4. Payload Structure

Use consistent payload structures that match the expected types:

```typescript
// Good - matches expected type
eventBus.emit(Events.DataChanged, [
  {
    referenceId: 'ref-123',
    oldValue: 100,
    newValue: 200,
    timestamp: Date.now(),
    source: 'manual'
  }
]);

// Bad - unexpected structure
eventBus.emit(Events.DataChanged, { foo: 'bar' });
```

## Singleton Pattern

The `EventBus` is exported as a singleton instance:

```typescript
export const eventBus = new EventBus();
```

This ensures all modules use the same event bus instance, enabling global event communication.

## Custom Events

While the `Events` enum covers common events, you can extend it for custom events:

```typescript
// In your module file
export const CustomEvents = {
  MyCustomEvent = 'my:custom:event'
} as const;

// Use it
eventBus.on(CustomEvents.MyCustomEvent, (data) => {
  console.log('Custom event received:', data);
});

eventBus.emit(CustomEvents.MyCustomEvent, { foo: 'bar' });
```

Note: You'll need to cast the event when using the EventBus API:

```typescript
eventBus.on(CustomEvents.MyCustomEvent as any, handler);
eventBus.emit(CustomEvents.MyCustomEvent as any, data);
```
