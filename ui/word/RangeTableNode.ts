import { Node, mergeAttributes } from '@tiptap/core';
import { Table } from '@tiptap/extension-table';
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
