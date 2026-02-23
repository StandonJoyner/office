'use client';

import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { TableRow, TableCell, TableHeader } from '@tiptap/extension-table';
import { ReferenceNode } from './ReferenceNode';
import { RangeTableNode } from './RangeTableNode';
import { DataReference } from '../../core/types';
import { eventBus, Events } from '../../core/eventBus';
import { useDataSync } from '../hooks/useDataSync';
import { useEditorDataSync } from '../hooks/useEditorDataSync';
import { ReferenceDetailsPanel } from '../common/ReferenceDetailsPanel';
import { useState, useEffect, forwardRef, useImperativeHandle, useCallback } from 'react';
import type { DataLinkManager } from '../../core/editors/DataLinkManager';
import type { DataSyncEngine } from '../../core/sync/DataSyncEngine';

interface WordEditorProps {
  content?: any;
  onChange?: (content: any) => void;
  onReferenceInserted?: (reference: DataReference) => void;
  /** Called when "+ Reference" is clicked. Parent should insert reference from selected Excel cell. */
  onInsertReferenceRequest?: () => void;
  /** Called when "+ Table Reference" is clicked. Parent should insert table reference from selected Excel range. */
  onInsertTableReferenceRequest?: () => void;
  /** When provided (e.g. from integrated page), use this documentId for Refresh All and same manager/syncEngine. */
  documentId?: string;
  /** When provided, use this manager instead of useDataSync internal one (must have setExcelWorkbook for resolve). */
  dataLinkManager?: DataLinkManager | null;
  /** When provided with dataLinkManager, used for Refresh All / refresh single. */
  syncEngine?: DataSyncEngine | null;
  /** When provided, called on Refresh All click so parent can use latest refs. Overrides internal refresh. */
  onRefreshAll?: () => Promise<void>;
  /** Optional ref to the refresh function; when set, called on Refresh All so parent always uses latest refs. */
  onRefreshAllRef?: React.MutableRefObject<(() => Promise<void>) | null>;
}

/**
 * Ref type for WordEditor component
 */
export interface WordEditorRefType {
  /**
   * Insert a data reference at the current cursor position
   */
  insertReference: (reference: DataReference) => void;
  /**
   * Insert a range reference table at the current cursor position
   */
  insertRangeTable: (reference: DataReference) => void;
  /**
   * Get the TipTap editor instance
   */
  getEditor: () => ReturnType<typeof useEditor> | null;
  /**
   * Get all reference IDs in the document
   */
  getReferenceIds: () => string[];
}

/**
 * Word Editor Component
 *
 * A rich text editor with data reference support and sync integration.
 */
export const WordEditor = forwardRef<WordEditorRefType, WordEditorProps>(({
  content,
  onChange,
  onReferenceInserted,
  onInsertReferenceRequest,
  onInsertTableReferenceRequest,
  documentId: documentIdProp,
  dataLinkManager: dataLinkManagerProp,
  syncEngine: syncEngineProp,
  onRefreshAll: onRefreshAllProp,
  onRefreshAllRef: onRefreshAllRefProp,
}, forwardedRef) => {
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [selectedReferenceId, setSelectedReferenceId] = useState<string | null>(null);

  const internalSync = useDataSync(documentIdProp ?? 'current-doc');
  const documentId = documentIdProp ?? 'current-doc';
  const dataLinkManager = dataLinkManagerProp ?? internalSync.dataLinkManager;
  const syncEngine = syncEngineProp ?? internalSync.syncEngine;

  const refreshAll = useCallback(async () => {
    const fn = onRefreshAllRefProp?.current ?? onRefreshAllProp;
    if (fn) {
      await fn();
      return;
    }
    if (syncEngine && documentId) await syncEngine.refreshDocument(documentId);
  }, [onRefreshAllRefProp, onRefreshAllProp, syncEngine, documentId]);

  const refreshReference = useCallback(async (refId: string) => {
    if (syncEngine) await syncEngine.refreshReference(refId);
  }, [syncEngine]);

  // Listen for reference click events
  useEffect(() => {
    const handleReferenceClicked = (data: any) => {
      if (data.referenceId) {
        setSelectedReferenceId(data.referenceId);
      }
    };

    eventBus.on('reference:clicked' as any, handleReferenceClicked);

    return () => {
      eventBus.off('reference:clicked' as any, handleReferenceClicked);
    };
  }, []);

  // Initialize editor
  const editor = useEditor({
    extensions: [StarterKit, TableRow, TableCell, TableHeader, ReferenceNode, RangeTableNode],
    content,
    editorProps: {
      attributes: {
        class: 'prose max-w-none focus:outline-none min-h-[500px] p-4',
      },
    },
    immediatelyRender: false,
    onUpdate: ({ editor }) => {
      const json = editor.getJSON();
      onChange?.(json);
    },
  });

  // Sync editor with data changes
  useEditorDataSync(editor);

  // Expose methods via ref for parent components
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

  const handleInsertReference = (reference: DataReference) => {
    if (!editor) return;

    // Store the reference under reference.id so document node refId matches manager
    if (dataLinkManager) {
      dataLinkManager.createReferenceWithId(reference.id, reference.source, reference.target.documentId);
      dataLinkManager.updateReference(reference.id, {
        display: reference.display,
        state: reference.state,
      });
    }

    // Insert the reference at cursor position
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'reference',
        attrs: {
          refId: reference.id,
          state: reference.state,
          value: reference.display.value,
          expression: reference.display.expression,
          format: reference.display.format,
          tooltip: reference.display.tooltip,
        },
      })
      .run();

    // Emit event for reference creation
    eventBus.emit(Events.ReferenceCreated, { reference });
    onReferenceInserted?.(reference);
  };

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

    const rowCount = tableData.length;
    const colCount = tableData[0]?.length || 0;
    if (rowCount === 0 || colCount === 0) return;

    // Parse source info for attributes
    const sourceInfo = {
      fileId: reference.source.fileId,
      sheetId: reference.source.sheetId,
      fileName: reference.source.fileName,
      range: reference.source.range,
    };

    // Build table content using Tiptap's table structure
    const tableContent = {
      type: 'rangeTable',
      attrs: {
        refId: reference.id,
        syncMode: reference.display.tableMeta?.syncMode || 'manual',
        sourceInfo: JSON.stringify(sourceInfo),
      },
      content: tableData.map((row) => ({
        type: 'tableRow',
        content: row.map((cell) => {
          const cellText = cell != null && cell !== '' ? String(cell) : '\u00A0';
          return {
            type: 'tableCell',
            content: [{ type: 'paragraph', content: [{ type: 'text', text: cellText }] }],
          };
        }),
      })),
    };

    // Insert table at cursor position
    editor
      .chain()
      .focus()
      .insertContent(tableContent)
      .run();

    // Emit event for reference creation
    eventBus.emit(Events.ReferenceCreated, { reference });
    onReferenceInserted?.(reference);
  };

  /**
   * Handle refresh all references
   */
  const handleRefreshAll = async () => {
    if (isRefreshing) return;
    setIsRefreshing(true);
    try {
      await refreshAll();
    } finally {
      setIsRefreshing(false);
    }
  };

  if (!editor) {
    return (
      <div className="flex items-center justify-center h-96 border border-slate-200 rounded-lg">
        <div className="text-slate-500">Loading editor...</div>
      </div>
    );
  }

  return (
    <div className="border border-slate-200 rounded-lg overflow-hidden">
      {/* Toolbar */}
      <div className="flex items-center gap-1 p-2 border-b border-slate-200 bg-slate-50">
        <TextFormatButtons editor={editor} />
        <div className="w-px h-6 bg-slate-300 mx-2" />
        <BlockFormatButtons editor={editor} />
        <div className="w-px h-6 bg-slate-300 mx-2" />

        {/* Data Sync Buttons */}
        <button
          type="button"
          onClick={handleRefreshAll}
          disabled={isRefreshing}
          className={`flex items-center gap-1.5 px-3 py-1.5 text-sm rounded ${
            isRefreshing
              ? 'bg-slate-300 text-slate-500 cursor-not-allowed'
              : 'bg-green-600 text-white hover:bg-green-700'
          }`}
          title="Refresh all data references"
        >
          {isRefreshing ? (
            <>
              <svg
                className="animate-spin w-4 h-4"
                fill="none"
                viewBox="0 0 24 24"
              >
                <circle
                  className="opacity-25"
                  cx="12"
                  cy="12"
                  r="10"
                  stroke="currentColor"
                  strokeWidth="4"
                />
                <path
                  className="opacity-75"
                  fill="currentColor"
                  d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                />
              </svg>
              <span>Refreshing...</span>
            </>
          ) : (
            <>
              <svg
                className="w-4 h-4"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
                />
              </svg>
              <span>Refresh All</span>
            </>
          )}
        </button>

        <div className="w-px h-6 bg-slate-300 mx-2" />

        <button
          type="button"
          onClick={() => {
            if (onInsertReferenceRequest) {
              // Parent (integrated page) inserts reference from selected Excel cell
              onInsertReferenceRequest();
            } else {
              // Fallback: insert test reference when no Excel context
              const testReference: DataReference = {
                id: crypto.randomUUID(),
                type: 'cell',
                source: {
                  fileId: 'test-file',
                  fileName: 'Test.xlsx',
                  sheetId: 'sheet1',
                  sheetName: 'Sheet1',
                  range: { startRow: 0, startCol: 1 },
                  isFormula: false,
                },
                target: {
                  documentId: 'test-doc',
                  nodeId: crypto.randomUUID(),
                },
                display: {
                  format: 'value',
                  value: '1,234.56',
                  expression: '[Test.xlsx!Sheet1!B1]',
                  tooltip: 'Source: Test.xlsx, Sheet1, B1',
                },
                state: 'active',
                history: [],
              };
              handleInsertReference(testReference);
            }
          }}
          className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700"
          title={onInsertReferenceRequest ? 'Insert reference from selected Excel cell' : 'Insert test reference'}
        >
          + Reference
        </button>

        <button
          type="button"
          onClick={() => {
            if (onInsertTableReferenceRequest) {
              onInsertTableReferenceRequest();
            } else {
              // Fallback: insert test table reference
              const testTableReference = {
                id: crypto.randomUUID(),
                type: 'range' as const,
                source: {
                  fileId: 'test-file',
                  fileName: 'Test.xlsx',
                  sheetId: 'sheet1',
                  sheetName: 'Sheet1',
                  range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 },
                  isFormula: false,
                },
                target: {
                  documentId: 'test-doc',
                  nodeId: crypto.randomUUID(),
                },
                display: {
                  format: 'value',
                  value: [
                    ['A1', 'B1', 'C1'],
                    ['A2', 'B2', 'C2'],
                    ['A3', 'B3', 'C3'],
                  ],
                  expression: '[Test.xlsx!Sheet1!A1:C3]',
                  tooltip: 'Source: Test.xlsx, Sheet1, A1:C3',
                  tableMeta: {
                    syncMode: 'manual' as const,
                    rowCount: 3,
                    colCount: 3,
                    preserveFormatting: false,
                  },
                },
                state: 'active' as const,
                history: [],
              };
              handleInsertRangeTable(testTableReference);
            }
          }}
          className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700"
          title={onInsertTableReferenceRequest ? 'Insert table reference from selected Excel range' : 'Insert test table reference'}
        >
          + Table Reference
        </button>
      </div>

      {/* Editor Content */}
      <EditorContent editor={editor} />

      {/* Reference Details Panel */}
      {selectedReferenceId && (
        <div className="border-t border-slate-200">
          <ReferenceDetailsPanel
            referenceId={selectedReferenceId}
            dataLinkManager={dataLinkManager}
            onClose={() => setSelectedReferenceId(null)}
          />
        </div>
      )}
    </div>
  );
});

WordEditor.displayName = 'WordEditor';

/**
 * Text format buttons (bold, italic, etc.)
 */
function TextFormatButtons({ editor }: { editor: any }) {
  const buttons = [
    { name: 'Bold', action: () => editor.chain().focus().toggleBold().run(), active: () => editor.isActive('bold'), icon: 'B' },
    { name: 'Italic', action: () => editor.chain().focus().toggleItalic().run(), active: () => editor.isActive('italic'), icon: 'I' },
    { name: 'Strike', action: () => editor.chain().focus().toggleStrike().run(), active: () => editor.isActive('strike'), icon: 'S' },
    { name: 'Code', action: () => editor.chain().focus().toggleCode().run(), active: () => editor.isActive('code'), icon: '<>' },
  ];

  return (
    <>
      {buttons.map((button) => (
        <button
          key={button.name}
          type="button"
          onClick={button.action}
          className={`p-2 rounded ${button.active() ? 'bg-blue-100 text-blue-700' : 'hover:bg-slate-100'}`}
          title={button.name}
        >
          <span className="font-bold">{button.icon}</span>
        </button>
      ))}
    </>
  );
}

/**
 * Block format buttons (headings, lists, etc.)
 */
function BlockFormatButtons({ editor }: { editor: any }) {
  const buttons = [
    { name: 'H1', action: () => editor.chain().focus().toggleHeading({ level: 1 }).run(), active: () => editor.isActive('heading', { level: 1 }) },
    { name: 'H2', action: () => editor.chain().focus().toggleHeading({ level: 2 }).run(), active: () => editor.isActive('heading', { level: 2 }) },
    { name: 'H3', action: () => editor.chain().focus().toggleHeading({ level: 3 }).run(), active: () => editor.isActive('heading', { level: 3 }) },
    { name: 'Bullet', action: () => editor.chain().focus().toggleBulletList().run(), active: () => editor.isActive('bulletList') },
    { name: 'Ordered', action: () => editor.chain().focus().toggleOrderedList().run(), active: () => editor.isActive('orderedList') },
  ];

  return (
    <>
      {buttons.map((button) => (
        <button
          key={button.name}
          type="button"
          onClick={button.action}
          className={`px-2 py-1 text-sm rounded ${button.active() ? 'bg-blue-100 text-blue-700' : 'hover:bg-slate-100'}`}
          title={button.name}
        >
          {button.name}
        </button>
      ))}
    </>
  );
}
