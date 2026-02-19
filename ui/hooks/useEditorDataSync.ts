'use client';

import { useEffect } from 'react';
import { Editor } from '@tiptap/react';
import { eventBus, Events } from '@/core/eventBus';
import { DataChange, DataReference } from '@/core/types';

/**
 * Editor Data Sync Hook
 *
 * Syncs TipTap editor with DataSyncEngine changes.
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
        }
      });
    };

    // Handle reference updates (new format - single reference event)
    const handleReferenceUpdated = (data: any) => {
      if (!editor) return;

      const { referenceId, newState } = data;

      // Find and update the reference node at the correct position (set selection then update)
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === 'reference' && node.attrs.refId === referenceId) {
          const newValue = newState?.display?.value ?? node.attrs.value;
          const state = newState?.state ?? node.attrs.state;

          editor.chain().setNodeSelection(pos).updateAttributes('reference', {
            refId: referenceId,
            value: newValue,
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
   * Find all reference IDs in the editor
   */
  const findReferenceIds = (): string[] => {
    if (!editor) return [];

    const refIds: string[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'reference' && node.attrs.refId) {
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
