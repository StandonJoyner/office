import { Node, mergeAttributes } from '@tiptap/core';
import { Table } from '@tiptap/extension-table';

export const RangeTableNode = Table.extend({
  name: 'rangeTable',

  content: 'tableRow+',

  group: 'block',

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
      // Render rows inside tbody
      ['tbody', 0],
    ];
  },
});
