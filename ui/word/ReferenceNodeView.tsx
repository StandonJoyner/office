'use client';

import { NodeViewWrapper, NodeViewProps } from '@tiptap/react';
import { useState } from 'react';
import { ReferenceState } from '../../core/types';
import { eventBus, Events } from '../../core/eventBus';

/**
 * State colors for reference display
 */
const STATE_STYLES: Record<
  ReferenceState,
  { bg: string; border: string; icon: string }
> = {
  active: {
    bg: 'bg-blue-50',
    border: 'border-blue-200',
    icon: '🔗',
  },
  stale: {
    bg: 'bg-yellow-50',
    border: 'border-yellow-200',
    icon: '⏳',
  },
  broken: {
    bg: 'bg-red-50',
    border: 'border-red-200',
    icon: '❌',
  },
  conflict: {
    bg: 'bg-orange-50',
    border: 'border-orange-200',
    icon: '⚠️',
  },
};

export function ReferenceNodeView(props: NodeViewProps) {
  const { node, updateAttributes, deleteNode, selected } = props;
  const attrs = node.attrs as {
    refId: string | null;
    state: ReferenceState;
    value: any;
    expression: string;
    format: 'value' | 'expression' | 'mixed';
    tooltip: string;
  };

  const [hovered, setHovered] = useState(false);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const state = attrs.state || 'active';
  const style = STATE_STYLES[state];

  const displayValue = (() => {
    switch (attrs.format) {
      case 'expression':
        return attrs.expression;
      case 'mixed':
        return `${attrs.value} (${attrs.expression})`;
      case 'value':
      default:
        return attrs.value;
    }
  })();

  const handleClick = () => {
    // Emit event to show reference details panel
    if (attrs.refId) {
      eventBus.emit('reference:clicked' as any, {
        referenceId: attrs.refId,
      });
    }
  };

  const handleRefresh = async (e: React.MouseEvent) => {
    e.stopPropagation();
    if (!attrs.refId || isRefreshing) return;

    setIsRefreshing(true);
    try {
      // Emit refresh event - will be handled by DataSyncEngine
      eventBus.emit(Events.ReferenceUpdated as any, {
        referenceId: attrs.refId,
        action: 'refresh',
      });

      // Simulate refresh - in production this would trigger the actual sync
      // For demo, update to a random new value
      const mockValues = ['1,234.56', '5,678.90', '3,456.78', '9,876.54', '2,345.67'];
      const newValue = mockValues[Math.floor(Math.random() * mockValues.length)];

      updateAttributes({
        value: newValue,
        state: 'active' as ReferenceState,
      });
    } catch (error) {
      console.error('Error refreshing reference:', error);
      updateAttributes({
        state: 'broken' as ReferenceState,
      });
    } finally {
      setIsRefreshing(false);
    }
  };

  const handleDelete = (e: React.MouseEvent) => {
    e.stopPropagation();
    deleteNode();
  };

  return (
    <NodeViewWrapper
      as="span"
      className="inline-flex items-center"
      contentEditable={false}
    >
      <span
        className={`
          inline-flex items-center gap-1 px-2 py-0.5 rounded text-sm
          ${style.bg} ${style.border} border
          cursor-pointer transition-all select-none
          ${selected ? 'ring-2 ring-blue-400 shadow-sm' : 'shadow-xs'}
          ${hovered ? `${style.border.replace('border-', 'hover:border-')} shadow-md` : ''}
          ${isRefreshing ? 'opacity-70' : ''}
        `}
        onClick={handleClick}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        title={attrs.tooltip || `Reference: ${attrs.refId}`}
      >
        <span className="text-xs">{style.icon}</span>
        <span className="font-medium">{displayValue}</span>

        {hovered && (
          <span className="flex items-center gap-1 ml-1">
            <button
              onClick={handleRefresh}
              disabled={isRefreshing}
              className={`p-0.5 hover:bg-white/50 rounded ${isRefreshing ? 'cursor-wait opacity-60' : ''}`}
              title="Refresh reference"
            >
              {isRefreshing ? (
                <svg
                  className="animate-spin w-3 h-3"
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
              ) : (
                <svg
                  className="w-3 h-3"
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
              )}
            </button>
            <button
              onClick={handleDelete}
              className="p-0.5 hover:bg-white/50 rounded hover:text-red-600 transition-colors"
              title="Remove reference"
            >
              <svg
                className="w-3 h-3"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M6 18L18 6M6 6l12 12"
                />
              </svg>
            </button>
          </span>
        )}
      </span>
    </NodeViewWrapper>
  );
}
