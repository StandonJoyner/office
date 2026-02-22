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
   * Update a range table with new data
   *
   * @param editor - The TipTap editor instance
   * @param pos - Position of the rangeTable node in the document
   * @param tableData - 2D array of new cell values
   */
  function updateRangeTable(editor: Editor, pos: number, tableData: any[][]): void {
    if (!tableData || !Array.isArray(tableData)) return;

    // Build new table rows from the 2D array
    const newRows = tableData.map((row) => {
      const cells = row.map((cell) => {
        return {
          type: 'tableCell',
          content: [{ type: 'paragraph', content: [{ type: 'text', text: String(cell ?? '') }] }],
        };
      });
      return { type: 'tableRow', content: cells };
    });

    // Get the current table node to preserve its attributes
    const node = editor.state.doc.nodeAt(pos);
    if (!node) return;

    const currentAttrs = { ...node.attrs };

    // Replace the table content with new rows while preserving attributes
    editor
      .chain()
      .setNodeSelection(pos)
      .insertContentAt(
        { from: pos, to: pos + node.nodeSize },
        {
          type: 'rangeTable',
          attrs: {
            ...currentAttrs,
            state: 'active',
          },
          content: newRows,
        }
      )
      .run();
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
