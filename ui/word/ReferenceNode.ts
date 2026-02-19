import { Node, mergeAttributes } from '@tiptap/core';
import { ReactNodeViewRenderer } from '@tiptap/react';
import { ReferenceNodeView } from './ReferenceNodeView';
import { DataReference, ReferenceState } from '../../core/types';

/**
 * Tiptap Node for Data Reference
 *
 * This is an inline atom node that represents a data reference to an Excel cell.
 * It can display the reference value, expression, or both.
 */
export const ReferenceNode = Node.create({
  name: 'reference',

  // Node is inline (can be part of a paragraph)
  group: 'inline',

  // Node is atomic (cannot have children)
  atom: true,

  // Make it inline (not a block element)
  inline: true,

  // Make node selectable
  selectable: true,

  // Make node draggable
  draggable: true,

  addAttributes() {
    return {
      refId: {
        default: null,
        parseHTML: (element) => element.getAttribute('data-ref-id'),
        renderHTML: (attributes) => {
          if (!attributes.refId) return {};
          return { 'data-ref-id': attributes.refId };
        },
      },
      state: {
        default: 'active' as ReferenceState,
        parseHTML: (element) => {
          const state = element.getAttribute('data-state') as ReferenceState;
          return state || 'active';
        },
        renderHTML: (attributes) => {
          return { 'data-state': attributes.state || 'active' };
        },
      },
      value: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-value'),
        renderHTML: (attributes) => {
          if (attributes.value === undefined || attributes.value === null) return {};
          return { 'data-value': String(attributes.value) };
        },
      },
      expression: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-expression'),
        renderHTML: (attributes) => {
          if (!attributes.expression) return {};
          return { 'data-expression': attributes.expression };
        },
      },
      format: {
        default: 'value',
        parseHTML: (element) =>
          element.getAttribute('data-format') || 'value',
        renderHTML: (attributes) => {
          return { 'data-format': attributes.format || 'value' };
        },
      },
      tooltip: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-tooltip'),
        renderHTML: (attributes) => {
          if (!attributes.tooltip) return {};
          return { 'data-tooltip': attributes.tooltip };
        },
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'span[data-type="reference"]',
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(
        {
          'data-type': 'reference',
          class: 'data-reference',
        },
        HTMLAttributes
      ),
    ];
  },

  addNodeView() {
    return ReactNodeViewRenderer(ReferenceNodeView);
  },
});
