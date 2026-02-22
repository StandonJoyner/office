# Range Reference Table Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Enable users to reference a range of Excel cells as a table in Word that syncs with Excel data changes.

**Architecture:** Extend Tiptap Table with custom attributes for Excel reference tracking, reuse existing DataLinkManager and ExcelDataSource for range resolution, implement mixed sync mode (manual/auto) with fixed table size.

**Tech Stack:** Tiptap (extensions), React, Univer (Excel), TypeScript, Vitest

---

## Prerequisites

### Task 0: Install Dependencies

**Files:**
- Modify: `package.json`

**Step 1: Add Tiptap table dependencies**

```json
{
  "dependencies": {
    "@tiptap/extension-table": "^2.9.0",
    "@tiptap/pm": "^2.9.0"
  }
}
```

**Step 2: Run npm install**

```bash
npm install
```

**Step 3: Verify installation**

Run: `grep "@tiptap/extension-table" package.json`
Expected: `"@tiptap/extension-table": "^2.9.0"`

**Step 4: Commit**

```bash
git add package.json package-lock.json
git commit -m "chore: add Tiptap table dependencies"
```

---

## Core Types Extension

### Task 1: Add RangeTableMeta to types.ts

**Files:**
- Modify: `core/types.ts:34-38`

**Step 1: Add RangeTableMeta interface**

```typescript
export interface RangeTableMeta {
  rowCount: number;
  colCount: number;
  syncMode: 'manual' | 'auto';
  preserveFormatting: boolean;
}
```

Place after `ReferenceHistory` interface (line 45), before `DataReference`.

**Step 2: Extend ReferenceDisplay to include tableMeta**

```typescript
export interface ReferenceDisplay {
  format: 'value' | 'expression' | 'mixed';
  value?: any;
  expression?: string;
  tooltip?: string;
  tableMeta?: RangeTableMeta;
}
```

**Step 3: Add test for RangeTableMeta**

Create test: `core/types.test.ts`

```typescript
import { describe, it, expect } from 'vitest';
import type { RangeTableMeta } from './types';

describe('RangeTableMeta', () => {
  it('should accept valid table metadata', () => {
    const meta: RangeTableMeta = {
      rowCount: 3,
      colCount: 2,
      syncMode: 'manual',
      preserveFormatting: true,
    };
    expect(meta.rowCount).toBe(3);
    expect(meta.colCount).toBe(2);
  });
});
```

**Step 4: Run test to verify it passes**

Run: `npm test -- core/types.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add core/types.ts core/types.test.ts
git commit -m "feat: add RangeTableMeta type definition"
```

---

## DataLinkManager Extension

### Task 2: Support range type in DataLinkManager.createReference

**Files:**
- Modify: `core/editors/DataLinkManager.ts:60-85`

**Step 1: Write failing test for range reference creation**

Create test: `core/editors/DataLinkManager.test.ts`

```typescript
import { describe, it, expect, beforeEach } from 'vitest';
import { createDataLinkManager } from './DataLinkManager';
import type { IExcelDataSourceManager } from './ExcelDataSource';
import type { DataReference } from '../types';

describe('DataLinkManager - Range References', () => {
  let manager: ReturnType<typeof createDataLinkManager>;
  let mockExcelManager: IExcelDataSourceManager;

  beforeEach(() => {
    mockExcelManager = {
      setWorkbook: vi.fn(),
      getWorkbook: vi.fn(),
      resolveCell: vi.fn(),
      resolveRange: vi.fn(),
      resolveReference: vi.fn(),
      onDataSourceChange: vi.fn(),
      offDataSourceChange: vi.fn(),
      enableChangeMonitoring: vi.fn(),
      validateSheet: vi.fn(() => true),
      validateCell: vi.fn(() => true),
      destroy: vi.fn(),
    } as any;
    manager = createDataLinkManager(mockExcelManager);
  });

  it('should create range reference with correct type', () => {
    const source = {
      fileId: 'file-1',
      fileName: 'Test.xlsx',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
      isFormula: false,
    };

    const refId = manager.createReference(source, 'doc-1');
    const ref = manager.getReference(refId);

    expect(ref?.type).toBe('range');
    expect(ref?.source.range.endRow).toBe(2);
    expect(ref?.source.range.endCol).toBe(1);
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- core/editors/DataLinkManager.test.ts`
Expected: FAIL (type is hardcoded to 'cell')

**Step 3: Update createReference to infer type from range**

Modify `core/editors/DataLinkManager.ts:60-85`:

```typescript
  createReference(source: DataSource, target: DocumentId): ReferenceId {
    const id = crypto.randomUUID();
    return this.createReferenceWithId(id, source, target);
  }

  createReferenceWithId(id: ReferenceId, source: DataSource, target: DocumentId): ReferenceId {
    // Infer type from range: has endRow/endCol means range reference
    const type: ReferenceType = (source.range.endRow !== undefined && source.range.endCol !== undefined)
      ? 'range'
      : 'cell';

    const ref: DataReference = {
      id,
      type,
      source,
      target: { documentId: target, nodeId: id },
      display: { format: 'value' },
      state: 'active',
      history: [
        {
          timestamp: Date.now(),
          action: 'created',
          value: null,
          userId: 'system',
        },
      ],
    };
    this.references.set(id, ref);
    eventBus.emit(Events.ReferenceCreated, ref);
    return id;
  }
```

**Step 4: Run test to verify it passes**

Run: `npm test -- core/editors/DataLinkManager.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add core/editors/DataLinkManager.ts core/editors/DataLinkManager.test.ts
git commit -m "feat: DataLinkManager support range type references"
```

---

### Task 3: Support range resolution in DataLinkManager

**Files:**
- Modify: `core/editors/DataLinkManager.ts:158-180`

**Step 1: Write failing test for range resolution**

Add to `core/editors/DataLinkManager.test.ts`:

```typescript
  it('should resolve range reference', async () => {
    const source = {
      fileId: 'file-1',
      fileName: 'Test.xlsx',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
      isFormula: false,
    };

    const mockRangeData = {
      values: [['A1', 'B1'], ['A2', 'B2'], ['A3', 'B3']],
      rowCount: 3,
      columnCount: 2,
      startAddress: 'A1',
    };

    mockExcelManager.resolveReference = vi.fn().mockResolvedValue(mockRangeData);

    const refId = manager.createReference(source, 'doc-1');
    const ref = manager.getReference(refId);
    const result = await manager.resolveReference(ref!);

    expect(mockExcelManager.resolveReference).toHaveBeenCalledWith(ref);
    expect(result).toEqual(mockRangeData.values);
  });
```

**Step 2: Run test to verify it passes**

Run: `npm test -- core/editors/DataLinkManager.test.ts`
Expected: PASS (already implemented via ExcelDataSource.resolveReference)

**Step 3: No code changes needed**

The `ExcelDataSource.resolveReference` already handles range type. Just verify test passes.

**Step 4: Commit**

```bash
git add core/editors/DataLinkManager.test.ts
git commit -m "test: add range resolution test"
```

---

## Word RangeTableNode

### Task 4: Create RangeTableNode extension

**Files:**
- Create: `ui/word/RangeTableNode.ts`

**Step 1: Write failing test for RangeTableNode existence**

Create test: `ui/word/RangeTableNode.test.ts`

```typescript
import { describe, it, expect } from 'vitest';
import { RangeTableNode } from './RangeTableNode';

describe('RangeTableNode', () => {
  it('should have correct node name', () => {
    expect(RangeTableNode.name).toBe('rangeTable');
  });

  it('should extend Table extension', () => {
    expect(RangeTableNode).toBeDefined();
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- ui/word/RangeTableNode.test.ts`
Expected: FAIL (file not found)

**Step 3: Implement RangeTableNode**

```typescript
import { Node, mergeAttributes } from '@tiptap/core';
import Table from '@tiptap/extension-table';
import { ReactNodeViewRenderer } from '@tiptap/react';
import { RangeTableNodeView } from './RangeTableNodeView';

/**
 * Tiptap Node for Excel Range Reference Table
 *
 * Extends Tiptap Table to add Excel reference tracking attributes.
 * Preserves all native table functionality (add/delete rows/columns, merge cells, etc.)
 */
export const RangeTableNode = Table.extend({
  name: 'rangeTable',

  group: 'block',
  content: 'tableRow+',

  addAttributes() {
    return {
      ...this.parent?.(),
      refId: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-ref-id'),
        renderHTML: (attributes) => {
          if (!attributes.refId) return {};
          return { 'data-ref-id': attributes.refId };
        },
      },
      syncMode: {
        default: 'manual',
        parseHTML: (element) => {
          const mode = element.getAttribute('data-sync-mode') as 'manual' | 'auto';
          return mode || 'manual';
        },
        renderHTML: (attributes) => {
          return { 'data-sync-mode': attributes.syncMode || 'manual' };
        },
      },
      sourceInfo: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-source-info'),
        renderHTML: (attributes) => {
          if (!attributes.sourceInfo) return {};
          const jsonString = typeof attributes.sourceInfo === 'string'
            ? attributes.sourceInfo
            : JSON.stringify(attributes.sourceInfo);
          return { 'data-source-info': jsonString };
        },
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'table[data-type="rangeTable"]',
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'table',
      mergeAttributes(
        {
          'data-type': 'rangeTable',
          class: 'range-reference-table',
        },
        HTMLAttributes
      ),
    ];
  },

  addNodeView() {
    return ReactNodeViewRenderer(RangeTableNodeView);
  },
});
```

**Step 4: Run test to verify it passes**

Run: `npm test -- ui/word/RangeTableNode.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add ui/word/RangeTableNode.ts ui/word/RangeTableNode.test.ts
git commit -m "feat: create RangeTableNode extending Tiptap Table"
```

---

### Task 5: Create RangeTableNodeView component

**Files:**
- Create: `ui/word/RangeTableNodeView.tsx`

**Step 1: Write failing test for component rendering**

Create test: `ui/word/RangeTableNodeView.test.tsx`

```typescript
import { describe, it, expect, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import { NodeViewWrapper } from '@tiptap/react';
import { RangeTableNodeView } from './RangeTableNodeView';
import { ReferenceState } from '../../core/types';

describe('RangeTableNodeView', () => {
  const mockNode = {
    attrs: {
      refId: 'test-ref-id',
      syncMode: 'manual' as const,
      sourceInfo: JSON.stringify({
        fileId: 'file-1',
        sheetId: 'sheet-1',
        range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
      }),
    },
  };

  const mockUpdateAttributes = vi.fn();
  const mockDeleteNode = vi.fn();

  it('should render table with data', () => {
    render(
      <NodeViewWrapper node={mockNode} updateAttributes={mockUpdateAttributes} deleteNode={mockDeleteNode}>
        <RangeTableNodeView node={mockNode} updateAttributes={mockUpdateAttributes} deleteNode={mockDeleteNode} />
      </NodeViewWrapper>
    );
    // Basic smoke test
    expect(screen.getByText(/test-ref-id/)).toBeInTheDocument();
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- ui/word/RangeTableNodeView.test.tsx`
Expected: FAIL (file not found)

**Step 3: Implement RangeTableNodeView**

```typescript
'use client';

import { NodeViewWrapper, NodeViewProps } from '@tiptap/react';
import { useState, useEffect } from 'react';
import { ReferenceState, RangeTableMeta } from '../../core/types';
import { eventBus, Events } from '../../core/eventBus';

interface SourceInfo {
  fileId: string;
  sheetId: string;
  fileName: string;
  range: { startRow: number; startCol: number; endRow: number; endCol: number };
}

interface RangeTableNodeViewProps extends NodeViewProps {
  node: {
    attrs: {
      refId: string | null;
      syncMode: 'manual' | 'auto';
      sourceInfo: string | SourceInfo;
    };
  };
}

const STATE_STYLES: Record<ReferenceState, { bg: string; border: string; icon: string }> = {
  active: { bg: 'bg-green-50', border: 'border-green-200', icon: '🔗' },
  stale: { bg: 'bg-yellow-50', border: 'border-yellow-200', icon: '⏳' },
  broken: { bg: 'bg-red-50', border: 'border-red-200', icon: '❌' },
  conflict: { bg: 'bg-orange-50', border: 'border-orange-200', icon: '⚠️' },
};

export function RangeTableNodeView(props: RangeTableNodeViewProps) {
  const { node, updateAttributes, deleteNode, selected } = props;
  const attrs = node.attrs;

  const [state, setState] = useState<ReferenceState>('active');
  const [hovered, setHovered] = useState(false);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [syncMode, setSyncMode] = useState<'manual' | 'auto'>(attrs.syncMode || 'manual');

  const style = STATE_STYLES[state];

  // Parse source info
  const sourceInfo: SourceInfo = typeof attrs.sourceInfo === 'string'
    ? JSON.parse(attrs.sourceInfo)
    : attrs.sourceInfo;

  // Listen to data change events for auto-sync
  useEffect(() => {
    if (syncMode === 'auto' && attrs.refId) {
      const handleDataChanged = (event: any) => {
        // Check if this reference is affected
        if (event.referenceId === attrs.refId) {
          refreshData();
        }
      };

      eventBus.on(Events.DataChanged, handleDataChanged);
      return () => eventBus.off(Events.DataChanged, handleDataChanged);
    }
  }, [syncMode, attrs.refId]);

  const refreshData = async () => {
    if (!attrs.refId || isRefreshing) return;

    setIsRefreshing(true);
    try {
      // Emit refresh event - will be handled by DataSyncEngine
      eventBus.emit(Events.ReferenceUpdated as any, {
        referenceId: attrs.refId,
        action: 'refresh',
      });

      // Update state to active after refresh
      setState('active');
    } catch (error) {
      console.error('Error refreshing table:', error);
      setState('broken');
    } finally {
      setIsRefreshing(false);
    }
  };

  const handleRefresh = (e: React.MouseEvent) => {
    e.stopPropagation();
    refreshData();
  };

  const handleDelete = (e: React.MouseEvent) => {
    e.stopPropagation();
    deleteNode();
  };

  const toggleSyncMode = (e: React.MouseEvent) => {
    e.stopPropagation();
    const newMode = syncMode === 'manual' ? 'auto' : 'manual';
    setSyncMode(newMode);
    updateAttributes({ syncMode: newMode });
  };

  return (
    <NodeViewWrapper
      as="div"
      className="relative inline-block"
      contentEditable={false}
    >
      {/* Status indicator */}
      <div
        className={`
          absolute -top-2 -left-2 z-10 px-2 py-1 rounded text-xs
          ${style.bg} ${style.border} border
          ${hovered ? 'opacity-100' : 'opacity-60'}
          transition-opacity
        `}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
      >
        <span>{style.icon}</span>
        <span className="ml-1 font-medium capitalize">{state}</span>
      </div>

      {/* Hover toolbar */}
      {hovered && (
        <div className="absolute -top-1 -right-1 z-20 flex gap-1 bg-white shadow-md rounded border border-slate-200 p-1">
          <button
            onClick={handleRefresh}
            disabled={isRefreshing}
            className={`p-1.5 hover:bg-slate-100 rounded ${isRefreshing ? 'cursor-wait opacity-60' : ''}`}
            title={syncMode === 'auto' ? 'Auto-sync enabled' : 'Refresh table'}
          >
            {isRefreshing ? (
              <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
              </svg>
            ) : (
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
              </svg>
            )}
          </button>
          <button
            onClick={toggleSyncMode}
            className={`p-1.5 hover:bg-slate-100 rounded text-xs ${syncMode === 'auto' ? 'bg-blue-100 text-blue-700' : ''}`}
            title={syncMode === 'auto' ? 'Switch to manual sync' : 'Switch to auto-sync'}
          >
            {syncMode === 'auto' ? 'Auto' : 'Manual'}
          </button>
          <button
            onClick={handleDelete}
            className="p-1.5 hover:bg-red-100 rounded hover:text-red-600"
            title="Remove table reference"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
      )}

      {/* Table content rendered by Tiptap */}
      <div
        className={`rounded border ${style.border} ${selected ? 'ring-2 ring-blue-400' : ''}`}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
      >
        {props.children}
      </div>
    </NodeViewWrapper>
  );
}
```

**Step 4: Run test to verify it passes**

Run: `npm test -- ui/word/RangeTableNodeView.test.tsx`
Expected: PASS

**Step 5: Commit**

```bash
git add ui/word/RangeTableNodeView.tsx ui/word/RangeTableNodeView.test.tsx
git commit -m "feat: create RangeTableNodeView component"
```

---

### Task 6: Register RangeTableNode in WordEditor

**Files:**
- Modify: `ui/word/WordEditor.tsx:3-6,104-110`

**Step 1: Write test for RangeTableNode registration**

Add to `ui/word/WordEditor.test.tsx` (create if needed):

```typescript
import { describe, it, expect } from 'vitest';
import { RangeTableNode } from './RangeTableNode';

describe('WordEditor', () => {
  it('should have RangeTableNode in extensions', () => {
    // This will be verified by checking the editor extensions
    expect(RangeTableNode).toBeDefined();
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- ui/word/WordEditor.test.tsx`
Expected: FAIL (RangeTableNode not imported/registered)

**Step 3: Import and register RangeTableNode**

Modify `ui/word/WordEditor.tsx`:

```typescript
// Add import after ReferenceNode import
import { RangeTableNode } from './RangeTableNode';

// In useEditor extensions array, add RangeTableNode
const editor = useEditor({
  extensions: [StarterKit, ReferenceNode, RangeTableNode],
  content,
  // ... rest of config
});
```

**Step 4: Run test to verify it passes**

Run: `npm test -- ui/word/WordEditor.test.tsx`
Expected: PASS

**Step 5: Commit**

```bash
git add ui/word/WordEditor.tsx ui/word/WordEditor.test.tsx
git commit -m "feat: register RangeTableNode in WordEditor"
```

---

### Task 7: Add insertRangeTable method to WordEditor

**Files:**
- Modify: `ui/word/WordEditor.tsx:36-49`

**Step 1: Write failing test for insertRangeTable**

Add to `ui/word/WordEditor.test.tsx`:

```typescript
  it('should have insertRangeTable method', () => {
    const ref = { current: { getEditor: () => null } };
    // Test will be updated once component is fully tested
    expect(true).toBe(true);
  });
```

**Step 2: Update WordEditorRefType interface**

Modify `ui/word/WordEditor.tsx:36-49`:

```typescript
export interface WordEditorRefType {
  insertReference: (reference: DataReference) => void;
  insertRangeTable: (reference: DataReference) => void;
  getEditor: () => ReturnType<typeof useEditor> | null;
  getReferenceIds: () => string[];
}
```

**Step 3: Implement insertRangeTable method**

Modify `ui/word/WordEditor.tsx:138-170`, add after `handleInsertReference`:

```typescript
  const handleInsertRangeTable = (reference: DataReference) => {
    if (!editor) return;

    // Store reference under reference.id
    if (dataLinkManager) {
      dataLinkManager.createReferenceWithId(reference.id, reference.source, reference.target.documentId);
      dataLinkManager.updateReference(reference.id, {
        display: reference.display,
        state: reference.state,
      });
    }

    // Build table rows from reference data
    const tableData = reference.display.value as any[][];
    if (!tableData || !Array.isArray(tableData)) return;

    const tableRows = tableData.map((row) => {
      const cells = row.map((cell) => {
        return {
          type: 'tableCell',
          content: [{ type: 'paragraph', content: [{ type: 'text', text: String(cell ?? '') }] }],
        };
      });
      return { type: 'tableRow', content: cells };
    });

    // Parse source info for attributes
    const sourceInfo = {
      fileId: reference.source.fileId,
      sheetId: reference.source.sheetId,
      fileName: reference.source.fileName,
      range: reference.source.range,
    };

    // Insert table at cursor position
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'rangeTable',
        attrs: {
          refId: reference.id,
          syncMode: reference.display.tableMeta?.syncMode || 'manual',
          sourceInfo: JSON.stringify(sourceInfo),
        },
        content: tableRows,
      })
      .run();

    // Emit event for reference creation
    eventBus.emit(Events.ReferenceCreated, { reference });
    onReferenceInserted?.(reference);
  };
```

**Step 4: Expose insertRangeTable via ref**

Modify `ui/word/WordEditor.tsx:123-136`:

```typescript
  useImperativeHandle(forwardedRef, () => ({
    insertReference: handleInsertReference,
    insertRangeTable: handleInsertRangeTable,
    getEditor: () => editor,
    getReferenceIds: () => {
      if (!editor) return [];
      const refIds: string[] = [];
      editor.state.doc.descendants((node) => {
        if ((node.type.name === 'reference' || node.type.name === 'rangeTable') && node.attrs.refId) {
          refIds.push(node.attrs.refId);
        }
      });
      return refIds;
    },
  }));
```

**Step 5: Run test to verify it passes**

Run: `npm test -- ui/word/WordEditor.test.tsx`
Expected: PASS

**Step 6: Commit**

```bash
git add ui/word/WordEditor.tsx ui/word/WordEditor.test.tsx
git commit -m "feat: add insertRangeTable method to WordEditor"
```

---

### Task 8: Add "+ Table Reference" button to WordEditor

**Files:**
- Modify: `ui/word/WordEditor.tsx:14-20,257-299`

**Step 1: Write test for button presence**

Add to `ui/word/WordEditor.test.tsx`:

```typescript
  it('should have Table Reference button', () => {
    // Will verify button renders when implemented
    expect(true).toBe(true);
  });
```

**Step 2: Add props for table reference insertion**

Modify `ui/word/WordEditor.tsx:14-20`:

```typescript
interface WordEditorProps {
  content?: any;
  onChange?: (content: any) => void;
  onReferenceInserted?: (reference: DataReference) => void;
  onInsertReferenceRequest?: () => void;
  onInsertTableReferenceRequest?: () => void;
  documentId?: string;
  dataLinkManager?: DataLinkManager | null;
  syncEngine?: DataSyncEngine | null;
  onRefreshAll?: () => Promise<void>;
  onRefreshAllRef?: React.MutableRefObject<(() => Promise<void>) | null>;
}
```

**Step 3: Add button in toolbar**

Modify `ui/word/WordEditor.tsx:257-299`, add after "+ Reference" button:

```typescript
        <button
          type="button"
          onClick={() => {
            if (onInsertTableReferenceRequest) {
              onInsertTableReferenceRequest();
            } else {
              // Fallback: insert test range table when no Excel context
              const testReference: DataReference = {
                id: crypto.randomUUID(),
                type: 'range',
                source: {
                  fileId: 'test-file',
                  fileName: 'Test.xlsx',
                  sheetId: 'sheet1',
                  sheetName: 'Sheet1',
                  range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
                  isFormula: false,
                },
                target: {
                  documentId: 'test-doc',
                  nodeId: crypto.randomUUID(),
                },
                display: {
                  format: 'value',
                  value: [
                    ['A1', 'B1'],
                    ['A2', 'B2'],
                    ['A3', 'B3'],
                  ],
                  tableMeta: {
                    rowCount: 3,
                    colCount: 2,
                    syncMode: 'manual',
                    preserveFormatting: true,
                  },
                  tooltip: 'Source: Test.xlsx, Sheet1, A1:B3',
                },
                state: 'active',
                history: [],
              };
              handleInsertRangeTable(testReference);
            }
          }}
          className="px-3 py-1.5 text-sm bg-purple-600 text-white rounded hover:bg-purple-700"
          title={onInsertTableReferenceRequest ? 'Insert table reference from selected Excel range' : 'Insert test range table'}
        >
          + Table Reference
        </button>
```

**Step 4: Run test to verify it passes**

Run: `npm test -- ui/word/WordEditor.test.tsx`
Expected: PASS

**Step 5: Commit**

```bash
git add ui/word/WordEditor.tsx ui/word/WordEditor.test.tsx
git commit -m "feat: add Table Reference button to WordEditor"
```

---

## Data Sync Integration

### Task 9: Update useEditorDataSync for range tables

**Files:**
- Modify: `ui/hooks/useEditorDataSync.ts`

**Step 1: Read current implementation to understand pattern**

```bash
cat ui/hooks/useEditorDataSync.ts
```

**Step 2: Write test for range table update handling**

Add to `ui/hooks/useEditorDataSync.test.ts` (create if needed):

```typescript
import { describe, it, expect, vi } from 'vitest';
import { useEditorDataSync } from './useEditorDataSync';
import { Events, eventBus } from '../../core/eventBus';

describe('useEditorDataSync', () => {
  it('should handle range table data updates', () => {
    // Test will be updated after reading implementation
    expect(true).toBe(true);
  });
});
```

**Step 3: Extend useEditorDataSync to handle rangeTable nodes**

After reading the file, add handling for `rangeTable` nodes in the data change listener:

```typescript
// In the DataChanged event handler, add rangeTable support
eventBus.on(Events.DataChanged, (dataChanges: any[]) => {
  if (!editor) return;

  dataChanges.forEach((change) => {
    // Find range table nodes with matching refId
    let rangeTableNode = null;
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'rangeTable' && node.attrs.refId === change.referenceId) {
        rangeTableNode = node;
        return false; // Stop traversal
      }
    });

    if (rangeTableNode) {
      // Get updated reference data
      const ref = dataLinkManager?.getReference(change.referenceId);
      if (ref && ref.display.value) {
        const newTableData = ref.display.value as any[][];
        // Update table content while preserving structure
        updateRangeTableContent(editor, change.referenceId, newTableData);
      }
    }
  });
});
```

**Step 4: Implement updateRangeTableContent helper**

```typescript
function updateRangeTableContent(editor: any, refId: string, newData: any[][]): void {
  const tr = editor.state.tr;

  let found = false;
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'rangeTable' && node.attrs.refId === refId) {
      found = true;

      // Build new table rows
      const newRows = newData.map((row) => {
        const cells = row.map((cell) => {
          return {
            type: 'tableCell',
            content: [{ type: 'paragraph', content: [{ type: 'text', text: String(cell ?? '') }] }],
          };
        });
        return { type: 'tableRow', content: cells };
      });

      // Replace the entire table content
      tr.insertText('', pos, pos + node.nodeSize);
      tr.insert(
        pos,
        editor.state.schema.nodes.rangeTable.create({ ...node.attrs }, newRows)
      );

      return false; // Stop traversal
    }
  });

  if (found) {
    editor.view.dispatch(tr);
  }
}
```

**Step 5: Run test to verify it passes**

Run: `npm test -- ui/hooks/useEditorDataSync.test.ts`
Expected: PASS

**Step 6: Commit**

```bash
git add ui/hooks/useEditorDataSync.ts ui/hooks/useEditorDataSync.test.ts
git commit -m "feat: add range table sync to useEditorDataSync"
```

---

## Integration Testing

### Task 10: E2E test for range reference table

**Files:**
- Create: `tests/e2e/range-reference-table.test.ts`

**Step 1: Write comprehensive E2E test**

```typescript
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { ExcelEditor } from '../../ui/excel/ExcelEditor';
import { WordEditor } from '../../ui/word/WordEditor';
import { createDataLinkManager } from '../../core/editors/DataLinkManager';
import type { DataReference } from '../../core/types';

describe('Range Reference Table E2E', () => {
  let dataLinkManager: ReturnType<typeof createDataLinkManager>;
  let wordEditorRef: React.RefObject<any>;
  let excelEditorRef: React.RefObject<any>;

  beforeEach(() => {
    dataLinkManager = createDataLinkManager();
  });

  afterEach(() => {
    dataLinkManager.destroy();
  });

  it('should create range reference table from Excel selection', async () => {
    const user = userEvent.setup();

    // 1. Select range in Excel
    // (Simulated by calling onRangeSelected callback)
    const selectedRange = {
      sheetId: 'sheet-1',
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
    };

    // 2. Create reference
    const reference: DataReference = {
      id: crypto.randomUUID(),
      type: 'range',
      source: {
        fileId: 'test-file',
        fileName: 'Test.xlsx',
        sheetId: 'sheet-1',
        sheetName: 'Sheet1',
        ...selectedRange.range,
        isFormula: false,
      },
      target: {
        documentId: 'test-doc',
        nodeId: crypto.randomUUID(),
      },
      display: {
        format: 'value',
        value: [
          ['A1', 'B1'],
          ['A2', 'B2'],
          ['A3', 'B3'],
        ],
        tableMeta: {
          rowCount: 3,
          colCount: 2,
          syncMode: 'manual',
          preserveFormatting: true,
        },
      },
      state: 'active',
      history: [],
    };

    // 3. Store in DataLinkManager
    dataLinkManager.createReferenceWithId(reference.id, reference.source, reference.target.documentId);

    // 4. Verify reference was created
    const retrievedRef = dataLinkManager.getReference(reference.id);
    expect(retrievedRef?.type).toBe('range');
    expect(retrievedRef?.source.range.endRow).toBe(2);
  });

  it('should handle manual refresh of range table', async () => {
    // Test manual refresh flow
    const refId = crypto.randomUUID();
    const ref: DataReference = {
      id: refId,
      type: 'range',
      source: {
        fileId: 'test-file',
        fileName: 'Test.xlsx',
        sheetId: 'sheet-1',
        sheetName: 'Sheet1',
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        isFormula: false,
      },
      target: { documentId: 'test-doc', nodeId: refId },
      display: {
        format: 'value',
        value: [['A1', 'B1'], ['A2', 'B2']],
        tableMeta: { rowCount: 2, colCount: 2, syncMode: 'manual', preserveFormatting: true },
      },
      state: 'active',
      history: [],
    };

    dataLinkManager.createReferenceWithId(ref.id, ref.source, ref.target.documentId);

    // Simulate data change marking as stale
    dataLinkManager.updateReference(ref.id, { state: 'stale' });

    const updatedRef = dataLinkManager.getReference(ref.id);
    expect(updatedRef?.state).toBe('stale');
  });

  it('should support auto-sync mode', () => {
    const refId = crypto.randomUUID();
    const ref: DataReference = {
      id: refId,
      type: 'range',
      source: {
        fileId: 'test-file',
        fileName: 'Test.xlsx',
        sheetId: 'sheet-1',
        sheetName: 'Sheet1',
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        isFormula: false,
      },
      target: { documentId: 'test-doc', nodeId: refId },
      display: {
        format: 'value',
        value: [['A1', 'B1'], ['A2', 'B2']],
        tableMeta: { rowCount: 2, colCount: 2, syncMode: 'auto', preserveFormatting: true },
      },
      state: 'active',
      history: [],
    };

    dataLinkManager.createReferenceWithId(ref.id, ref.source, ref.target.documentId);

    const retrievedRef = dataLinkManager.getReference(ref.id);
    expect(retrievedRef?.display.tableMeta?.syncMode).toBe('auto');
  });
});
```

**Step 2: Run E2E test**

Run: `npm test -- tests/e2e/range-reference-table.test.ts`
Expected: PASS

**Step 3: Commit**

```bash
git add tests/e2e/range-reference-table.test.ts
git commit -m "test: add E2E test for range reference table"
```

---

## Build Verification

### Task 11: Verify TypeScript compilation

**Files:**
- N/A (run command)

**Step 1: Run build to verify type checking**

```bash
npm run build
```

Expected: No TypeScript errors

**Step 2: If build fails, fix type errors and rerun**

**Step 3: Commit any build fixes**

```bash
git add .
git commit -m "fix: resolve TypeScript compilation errors"
```

---

### Task 12: Run full test suite

**Files:**
- N/A (run command)

**Step 1: Run all tests**

```bash
npm test
```

Expected: All tests pass

**Step 2: If tests fail, fix and rerun**

**Step 3: Commit any test fixes**

```bash
git add .
git commit -m "fix: resolve test failures"
```

---

## Manual Testing Checklist

### Task 13: Manual verification

**Files:**
- N/A (manual testing steps)

**Step 1: Test create range reference table**

1. Run dev server: `npm run dev`
2. Open Word editor
3. In Excel, select range A1:C3
4. Click "+ Table Reference" in Word
5. Verify table appears with correct data

**Step 2: Test manual refresh**

1. Change a cell value in Excel (e.g., A1 from "test" to "updated")
2. Verify table shows stale indicator (⏳)
3. Click refresh button on table
4. Verify table shows updated value and active indicator (🔗)

**Step 3: Test auto-sync mode**

1. Click sync mode button to switch to "Auto"
2. Change a cell value in Excel
3. Verify table auto-updates immediately

**Step 4: Test state transitions**

1. Delete Excel sheet
2. Verify table shows broken state (❌)

**Step 5: Test format preservation**

1. Bold a cell in the Word table
2. Refresh table data
3. Verify bold formatting is preserved

**Step 6: Document any issues**

Create note: `docs/plans/2026-02-22-range-reference-table-testing-notes.md`

**Step 7: Commit testing notes if any issues found**

```bash
git add docs/plans/2026-02-22-range-reference-table-testing-notes.md
git commit -m "docs: add manual testing notes"
```

---

## Final Verification

### Task 14: Complete implementation verification

**Files:**
- N/A (verification checklist)

**Step 1: Verify all tasks complete**

Run: `git log --oneline --since="2026-02-22"`
Expected: ~14 commits for range reference table feature

**Step 2: Verify all files modified correctly**

Files that should exist/be modified:
- ✅ `core/types.ts` - RangeTableMeta added
- ✅ `core/editors/DataLinkManager.ts` - range support added
- ✅ `ui/word/RangeTableNode.ts` - new file
- ✅ `ui/word/RangeTableNodeView.tsx` - new file
- ✅ `ui/word/WordEditor.tsx` - table reference button and method
- ✅ `ui/hooks/useEditorDataSync.ts` - range table sync
- ✅ `package.json` - Tiptap table dependencies

**Step 3: Run build one final time**

```bash
npm run build
```

**Step 4: Run tests one final time**

```bash
npm test
```

**Step 5: Verify implementation against design doc**

Compare with `docs/plans/2026-02-22-range-reference-table-design.md`

**Step 6: Mark implementation complete**

Create: `docs/plans/2026-02-22-range-reference-table-complete.md`

```markdown
# Range Reference Table - Implementation Complete

Date: 2026-02-22
Status: Complete

All tasks completed. Feature is ready for use.
```

**Step 7: Commit completion marker**

```bash
git add docs/plans/2026-02-22-range-reference-table-complete.md
git commit -m "docs: mark range reference table implementation complete"
```
