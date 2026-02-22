'use client';

import { NodeViewWrapper, NodeViewProps } from '@tiptap/react';
import { useState, useEffect, ReactNode } from 'react';
import { ReferenceState, RangeTableMeta } from '../../core/types';
import { eventBus, Events } from '../../core/eventBus';

interface SourceInfo {
  fileId: string;
  sheetId: string;
  fileName: string;
  range: { startRow: number; startCol: number; endRow: number; endCol: number };
}

interface RangeTableNodeAttrs {
  refId: string | null;
  syncMode: 'manual' | 'auto';
  sourceInfo: string | SourceInfo;
}

interface RangeTableNodeViewProps extends NodeViewProps {
  children?: ReactNode;
  /** Optional initial state for testing purposes. Not used in production. */
  initialState?: ReferenceState;
}

const STATE_STYLES: Record<ReferenceState, { bg: string; border: string; icon: string }> = {
  active: { bg: 'bg-green-50', border: 'border-green-200', icon: '🔗' },
  stale: { bg: 'bg-yellow-50', border: 'border-yellow-200', icon: '⏳' },
  broken: { bg: 'bg-red-50', border: 'border-red-200', icon: '❌' },
  conflict: { bg: 'bg-orange-50', border: 'border-orange-200', icon: '⚠️' },
};

export function RangeTableNodeView(props: RangeTableNodeViewProps) {
  const { node, updateAttributes, deleteNode, selected, initialState } = props;
  const attrs = node.attrs as RangeTableNodeAttrs;

  const [state, setState] = useState<ReferenceState>(initialState ?? 'active');
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
