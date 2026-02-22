# Range Reference Table Design

## Context

Single cell references have been implemented. Users now need the ability to reference a range of cells from Excel that inserts as a table in Word. The table data should sync with Excel cell changes.

## Requirements

1. **Sync Mode**: Mixed mode (default manual refresh, configurable for auto-sync)
2. **Table Size**: Fixed size - row/column count determined at creation, data changes only update cell values
3. **Header Handling**: No special header treatment - user selects which rows to include
4. **Insert Flow**: User selects range in Excel, then clicks "+ Table Reference" button in Word
5. **Table Operations**: Must have parity with Tiptap native table functionality (add/delete rows/columns, merge cells, styling)

## Architecture

```
Excel (range selection) → DataLinkManager (type='range' ref) → Word (RangeTableNode)
                           ↓
                   ExcelDataSource (resolveRange)
                           ↓
                   DataSyncEngine (sync logic)
                           ↓
                   EventBus (DataChanged events)
```

### Core Components

| Component | Purpose |
|-----------|---------|
| `RangeTableNode` | Extends Tiptap Table, adds reference attributes |
| `RangeTableNodeView` | React component for table rendering and interaction |
| Extended `DataLinkManager` | Supports creating and resolving 'range' type references |
| `ExcelDataSource.resolveRange()` | Already implemented - fetches 2D cell data |
| `DataSyncEngine` | Reused for sync logic (manual/auto modes) |

## Data Model Extensions

### Extended ReferenceDisplay

```typescript
interface RangeTableMeta {
  rowCount: number;
  colCount: number;
  syncMode: 'manual' | 'auto';
  preserveFormatting: boolean;
}

interface ReferenceDisplay {
  // Existing fields...
  tableMeta?: RangeTableMeta;  // New field for range references
}
```

### DataReference (type='range')

```typescript
{
  id: string;
  type: 'range';
  source: DataSource {
    fileId: string;
    sheetId: string;
    range: CellRange { startRow, startCol, endRow, endCol };
  };
  target: ReferenceTarget;
  display: {
    value: any[][];  // 2D array of cell values
    tableMeta: RangeTableMeta;
  };
  state: 'active' | 'stale' | 'broken' | 'conflict';
}
```

## Word Implementation

### 1. RangeTableNode (ui/word/RangeTableNode.ts)

Extends `@tiptap/extension-table`:

```typescript
const RangeTableNode = Table.extend({
  name: 'rangeTable',

  addAttributes() {
    return {
      ...this.parent?.(),
      refId: { default: null },
      syncMode: { default: 'manual' },
      sourceInfo: { default: null },  // JSON string of {fileId, sheetId, range}
    };
  },

  addNodeView() {
    return ReactNodeViewRenderer(RangeTableNodeView);
  },

  parseHTML() {
    return [
      {
        tag: 'table[data-type="rangeTable"]',
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return ['table', mergeAttributes({
      'data-type': 'rangeTable',
    }, HTMLAttributes)];
  },
});
```

### 2. RangeTableNodeView (ui/word/RangeTableNodeView.tsx)

Features:
- Renders table with Excel data
- Status indicator in top-left corner (active/stale/broken)
- Hover toolbar: Refresh, Delete, Toggle Sync Mode
- Listens to `Events.DataChanged` for auto-sync
- Preserves user formatting changes

```typescript
interface RangeTableNodeViewProps extends NodeViewProps {
  node.attrs: {
    refId: string;
    syncMode: 'manual' | 'auto';
    sourceInfo: string;  // JSON string
  };
}
```

### 3. WordEditor Extensions (ui/word/WordEditor.tsx)

Add:
- "+ Table Reference" button in toolbar
- `insertRangeTable(reference: DataReference)` method
- Register `RangeTableNode` extension

## Excel Integration

The `ExcelEditor` already has `onRangeSelected` callback:

```typescript
// Store selected range when user selects multiple cells
const [selectedRange, setSelectedRange] = useState<{
  sheetId: string;
  range: CellRange;
} | null>(null);

// Pass selected range when creating reference
<ExcelEditor onRangeSelected={setSelectedRange} />
```

## Data Sync Flow

### Create Reference

1. User selects range A1:C3 in Excel
2. Click "+ Table Reference" in Word
3. `DataLinkManager.createReference(source, docId)` with `type='range'`
4. `ExcelDataSource.resolveRange()` fetches `any[][]`
5. Insert `RangeTableNode` with initial data

### Manual Refresh

1. Excel data changes → `DataSourceEvent` emitted
2. `DataLinkManager` marks reference as `stale`
3. User clicks refresh button on table
4. `ExcelDataSource.resolveRange()` updates data
5. Update table cell values (preserve formatting)

### Auto Sync

1. Excel data changes → `DataSourceEvent` emitted
2. `DataSyncEngine` detects auto-sync enabled
3. Auto-trigger `refreshReference()`
4. Update table cell values (preserve formatting)

### State Visualization

| State | Visual Indicator | Behavior |
|-------|----------------|----------|
| active | 🔗 Green badge | Data is current |
| stale | ⏳ Yellow badge | Data changed, needs refresh |
| broken | ❌ Red badge | Source unavailable |
| conflict | ⚠️ Orange badge | Concurrent edit conflict |

## File Changes

| File | Type | Changes |
|------|------|---------|
| `ui/word/RangeTableNode.ts` | New | Extends Tiptap Table with reference attributes |
| `ui/word/RangeTableNodeView.tsx` | New | Table rendering and interaction |
| `ui/word/WordEditor.tsx` | Modify | Add table reference button, register RangeTableNode |
| `core/editors/DataLinkManager.ts` | Modify | Support range type references in create/resolve |
| `core/types.ts` | Modify | Add RangeTableMeta interface |
| `ui/hooks/useEditorDataSync.ts` | Modify | Handle table node updates from DataChanged events |
| `package.json` | Modify | Add `@tiptap/extension-table` and `@tiptap/pm` dependencies |

## Dependencies

Add to `package.json`:
```json
{
  "@tiptap/extension-table": "^2.9.0",
  "@tiptap/pm": "^2.9.0",
  "@tiptap/starter-kit": "^2.9.0"
}
```

## Testing

1. **Create table reference**: Select range, insert table, verify data matches
2. **Manual refresh**: Change Excel data, verify table shows stale, refresh updates
3. **Auto sync**: Set auto mode, verify table updates immediately
4. **State transitions**: Verify active → stale → active flow
5. **Format preservation**: Edit table formatting, refresh data, verify format preserved
6. **Error handling**: Delete Excel sheet, verify broken state

## Future Enhancements

- Configurable header row (first row as header)
- Dynamic table size option
- Table-level formatting templates
- Export referenced tables as standalone
- Support named ranges
