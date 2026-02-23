'use client';

import { useEffect } from 'react';
import { Editor } from '@tiptap/react';
import { eventBus, Events } from '@/core/eventBus';
import { DataChange, DataReference } from '@/core/types';

/**
 * Editor Data Sync Hook
 *
 * Syncs TipTap editor with DataSyncEngine changes.
 * Handles both single cell references and range table references.
 */
export function useEditorDataSync(editor: Editor | null) {
  useEffect(() => {
    if (!editor) return;

    // Handle data changes (legacy format - array of changes)
    const handleDataChanged = (changes: DataChange[]) => {
      if (!editor || changes.length === 0) return;

      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === 'reference') {
          const refId = node.attrs.refId;
          const change = changes.find(c => c.referenceId === refId);

          if (change) {
            editor.chain().setNodeSelection(pos).updateAttributes('reference', {
              refId,
              value: change.newValue,
              state: 'active',
            }).run();
          }
        } else if (node.type.name === 'rangeTable') {
          const refId = node.attrs.refId;
          const change = changes.find(c => c.referenceId === refId);

          if (change && Array.isArray(change.newValue)) {
            updateRangeTable(editor, pos, change.newValue);
          }
        }
      });
    };

    // Handle reference updates (new format - single reference event)
    const handleReferenceUpdated = (data: any) => {
      if (!editor) return;

      const { referenceId, newState } = data;

      // Find and update reference nodes
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === 'reference' && node.attrs.refId === referenceId) {
          const newValue = newState?.display?.value ?? node.attrs.value;
          const state = newState?.state ?? node.attrs.state;

          editor.chain().setNodeSelection(pos).updateAttributes('reference', {
            refId: referenceId,
            value: newValue,
            state,
          }).run();
        } else if (node.type.name === 'rangeTable' && node.attrs.refId === referenceId) {
          // Handle range table updates
          const tableData = newState?.display?.value;
          if (tableData && Array.isArray(tableData)) {
            updateRangeTable(editor, pos, tableData);
          }
          // Update table attributes (state)
          const state = newState?.state ?? node.attrs.state;
          editor.chain().setNodeSelection(pos).updateAttributes('rangeTable', {
            state,
          }).run();
        }
      });
    };

    // Listen for data changes
    eventBus.on(Events.DataChanged, handleDataChanged);
    eventBus.on(Events.ReferenceUpdated, handleReferenceUpdated);

    return () => {
      eventBus.off(Events.DataChanged, handleDataChanged);
      eventBus.off(Events.ReferenceUpdated, handleReferenceUpdated);
    };
  }, [editor]);

  /**
   * Update a range table with new data, preserving per-cell formatting (marks) in Word.
   * Each cell is treated as one block; only the text value is updated, existing marks are kept.
   */
  function updateRangeTable(editor: Editor, pos: number, tableData: any[][]): void {
    if (!tableData || !Array.isArray(tableData)) return;

    const node = editor.state.doc.nodeAt(pos);
    if (!node) return;

    const currentAttrs = { ...node.attrs };

    // Extract marks from the first text node in each cell (ProseMirror/Tiptap)
    const preservedMarks = getCellMarksFromTable(node);

    const newRows = tableData.map((row, rowIndex) => ({
      type: 'tableRow',
      content: row.map((cell, colIndex) => {
        const cellText = cell != null && cell !== '' ? String(cell) : '\u00A0';
        const marks = preservedMarks[rowIndex]?.[colIndex] ?? [];
        return {
          type: 'tableCell',
          content: [
            {
              type: 'paragraph',
              content: [{ type: 'text', text: cellText, marks }],
            },
          ],
        };
      }),
    }));

    editor
      .chain()
      .setNodeSelection(pos)
      .insertContentAt(
        { from: pos, to: pos + node.nodeSize },
        {
          type: 'rangeTable',
          attrs: { ...currentAttrs, state: 'active' },
          content: newRows,
        }
      )
      .run();
  }

  /**
   * Returns a 2D array [row][col] of mark specs for each cell (for insertContent).
   */
  function getCellMarksFromTable(tableNode: { forEach: (f: (node: unknown) => void) => void }): Record<string, unknown>[][] {
    const rows: Record<string, unknown>[][] = [];
    let rowIndex = 0;
    tableNode.forEach((rowNode: unknown) => {
      const r = rowNode as { type: { name: string }; forEach: (f: (node: unknown) => void) => void };
      if (r.type?.name !== 'tableRow') return;
      rows[rowIndex] = [];
      let colIndex = 0;
      r.forEach((cellNode: unknown) => {
        const c = cellNode as { type: { name: string }; forEach: (f: (node: unknown) => void) => void };
        if (c.type?.name !== 'tableCell') return;
        rows[rowIndex][colIndex] = getMarksFromCell(c);
        colIndex++;
      });
      rowIndex++;
    });
    return rows;
  }

  function getMarksFromCell(cellNode: { forEach: (f: (node: unknown) => void) => void }): Record<string, unknown>[] {
    let firstTextMarks: Record<string, unknown>[] = [];
    cellNode.forEach((blockNode: unknown) => {
      if (firstTextMarks.length > 0) return;
      const b = blockNode as { type: { name: string }; forEach: (f: (node: unknown) => void) => void };
      if (b.type?.name !== 'paragraph') return;
      b.forEach((inlineNode: unknown) => {
        const i = inlineNode as { type: { name: string }; marks?: readonly { type: { name: string }; attrs?: Record<string, unknown> }[] };
        if (i.type?.name === 'text' && i.marks?.length) {
          firstTextMarks = i.marks.map((m) =>
            (m as { toJSON?: () => Record<string, unknown> }).toJSON
              ? (m as { toJSON: () => Record<string, unknown> }).toJSON()
              : { type: m.type.name, ...(m.attrs && Object.keys(m.attrs).length ? { attrs: m.attrs } : {}) }
          );
        }
      });
    });
    return firstTextMarks;
  }

  /**
   * Find all reference IDs in the editor (including range table references)
   */
  const findReferenceIds = (): string[] => {
    if (!editor) return [];

    const refIds: string[] = [];
    editor.state.doc.descendants((node) => {
      if (
        (node.type.name === 'reference' || node.type.name === 'rangeTable') &&
        node.attrs.refId
      ) {
        refIds.push(node.attrs.refId);
      }
    });

    return refIds;
  };

  /**
   * Update a specific reference in the editor
   */
  const updateReference = (refId: string, newValue: any, state: string = 'active') => {
    if (!editor) return;

    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'reference' && node.attrs.refId === refId) {
        editor.commands.setNodeSelection(pos);
        editor.commands.updateAttributes('reference', {
          value: newValue,
          state,
        });
      }
    });
  };

  return {
    findReferenceIds,
    updateReference,
  };
}
