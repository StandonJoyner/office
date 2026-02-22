'use client';

import { NodeViewWrapper, NodeViewProps } from '@tiptap/react';

/**
 * React Node View Component for Range Reference Table
 *
 * This component wraps the Tiptap table node with additional UI for
 * Excel reference tracking (sync mode display, refresh controls, etc.)
 *
 * NOTE: This is a stub implementation that will be fully implemented in Task #6.
 */
export function RangeTableNodeView(props: NodeViewProps) {
  const { node } = props;

  return (
    <NodeViewWrapper as="div" className="range-table-wrapper">
      <table className="range-reference-table">
        {node.content.childCount > 0 && (
          <tbody>
            {node.content.map((child, i) => {
              // Render table rows - placeholder rendering
              return (
                <tr key={i}>
                  <td>Range table cell</td>
                </tr>
              );
            })}
          </tbody>
        )}
      </table>
    </NodeViewWrapper>
  );
}
