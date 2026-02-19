'use client';

import { useState, useMemo } from 'react';
import {
  TraceabilityService,
  createTraceabilityService,
  AuditLog,
  Diff,
  DiffChange,
  AuditFilters,
  AuditActionType,
} from '../../core/trace/TraceabilityService';
import type { DocumentSnapshot } from '../../core/types';
import { ReferenceState } from '../../core/types';

/**
 * Data Traceability Panel Component
 *
 * Displays version snapshots, change history, and dependency relationships.
 */
interface TraceabilityPanelProps {
  documentId?: string;
  referenceId?: string;
  traceabilityService: TraceabilityService;
  onReferenceClick?: (referenceId: string) => void;
}

type TabType = 'history' | 'snapshots' | 'dependencies' | 'audit';

export function TraceabilityPanel({
  documentId,
  referenceId,
  traceabilityService,
  onReferenceClick,
}: TraceabilityPanelProps) {
  const [activeTab, setActiveTab] = useState<TabType>('history');
  const [selectedSnapshotId, setSelectedSnapshotId] = useState<string | null>(null);
  const [compareSnapshotId, setCompareSnapshotId] = useState<string | null>(null);
  const [auditFilters, setAuditFilters] = useState<AuditFilters>({});
  const [searchQuery, setSearchQuery] = useState('');
  const [isPanelCollapsed, setIsPanelCollapsed] = useState(false);

  // Fetch data based on active tab
  const referenceHistory = useMemo(
    () => (referenceId ? traceabilityService.getReferenceHistory(referenceId) : []),
    [referenceId, traceabilityService]
  );

  const snapshots = useMemo(
    () => (documentId ? traceabilityService.listSnapshots(documentId) : []),
    [documentId, traceabilityService]
  );

  const dependencies = useMemo(
    () => (referenceId ? traceabilityService.getDependents(referenceId) : []),
    [referenceId, traceabilityService]
  );

  const auditLogs = useMemo(
    () => traceabilityService.getAuditLog(auditFilters),
    [auditFilters, traceabilityService]
  );

  // Snapshot comparison
  const snapshotDiff = useMemo(
    () =>
      selectedSnapshotId && compareSnapshotId
        ? traceabilityService.compareSnapshots(selectedSnapshotId, compareSnapshotId)
        : null,
    [selectedSnapshotId, compareSnapshotId, traceabilityService]
  );

  // Filter data based on search query
  const filteredHistory = useMemo(
    () =>
      referenceHistory.filter(
        (entry) =>
          !searchQuery ||
          entry.action.toLowerCase().includes(searchQuery.toLowerCase()) ||
          (entry.userId && entry.userId.toLowerCase().includes(searchQuery.toLowerCase()))
      ),
    [referenceHistory, searchQuery]
  );

  const filteredAuditLogs = useMemo(() => {
    let logs = auditLogs.slice(0, auditFilters.limit || 50);
    if (searchQuery) {
      logs = logs.filter(
        (log) =>
          log.action.toLowerCase().includes(searchQuery.toLowerCase()) ||
          log.userId.toLowerCase().includes(searchQuery.toLowerCase()) ||
          (log.metadata &&
            Object.values(log.metadata).some(
              (v) =>
                v && String(v).toLowerCase().includes(searchQuery.toLowerCase())
            ))
      );
    }
    return logs;
  }, [auditLogs, searchQuery, auditFilters.limit]);

  const tabs: Array<{ id: TabType; label: string; icon: string; count: number }> = [
    {
      id: 'history',
      label: 'History',
      icon: '📜',
      count: filteredHistory.length,
    },
    {
      id: 'snapshots',
      label: 'Snapshots',
      icon: '📸',
      count: snapshots.length,
    },
    {
      id: 'dependencies',
      label: 'Dependencies',
      icon: '🔗',
      count: dependencies.length,
    },
    {
      id: 'audit',
      label: 'Audit Log',
      icon: '📋',
      count: filteredAuditLogs.length,
    },
  ];

  // Create snapshot handler
  const handleCreateSnapshot = async () => {
    if (!documentId) return;
    await traceabilityService.createSnapshot(documentId);
    setActiveTab('snapshots');
  };

  // Restore snapshot handler
  const handleRestoreSnapshot = async (snapshotId: string) => {
    const snapshot = traceabilityService.getSnapshot(snapshotId);
    if (snapshot) {
      // Trigger document restore event
      console.log('Restore snapshot:', snapshot);
    }
  };

  // Delete snapshot handler
  const handleDeleteSnapshot = (snapshotId: string) => {
    traceabilityService.deleteSnapshot(snapshotId);
    if (selectedSnapshotId === snapshotId) {
      setSelectedSnapshotId(null);
    }
  };

  // Compare snapshots handler
  const handleCompareSnapshots = (snapshotId: string) => {
    if (selectedSnapshotId === snapshotId) {
      setCompareSnapshotId(null);
    } else {
      setCompareSnapshotId(snapshotId);
    }
  };

  const getStateStyle = (state: ReferenceState) => {
    switch (state) {
      case 'active':
        return {
          bg: 'bg-green-50',
          border: 'border-green-200',
          text: 'text-green-700',
          dot: 'bg-green-500',
        };
      case 'stale':
        return {
          bg: 'bg-yellow-50',
          border: 'border-yellow-200',
          text: 'text-yellow-700',
          dot: 'bg-yellow-500',
        };
      case 'broken':
        return {
          bg: 'bg-red-50',
          border: 'border-red-200',
          text: 'text-red-700',
          dot: 'bg-red-500',
        };
      case 'conflict':
        return {
          bg: 'bg-orange-50',
          border: 'border-orange-200',
          text: 'text-orange-700',
          dot: 'bg-orange-500',
        };
      default:
        return {
          bg: 'bg-slate-50',
          border: 'border-slate-200',
          text: 'text-slate-700',
          dot: 'bg-slate-500',
        };
    }
  };

  if (isPanelCollapsed) {
    return (
      <button
        onClick={() => setIsPanelCollapsed(false)}
        className="bg-white rounded-lg shadow-md border border-slate-200 p-3 hover:bg-slate-50 transition-colors"
        title="Expand traceability panel"
      >
        <div className="flex items-center gap-2">
          <span className="text-2xl">📜</span>
          <span className="font-medium text-slate-700">Data Traceability</span>
          <svg
            className="w-4 h-4 text-slate-500"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M19 9l-7 7-7-7"
            />
          </svg>
        </div>
      </button>
    );
  }

  return (
    <div className="bg-white rounded-lg shadow-lg border border-slate-200 overflow-hidden">
      {/* Header */}
      <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3 bg-slate-50">
        <div className="flex items-center gap-2">
          <span className="text-xl">📜</span>
          <h2 className="text-lg font-semibold text-slate-900">Data Traceability</h2>
        </div>
        <div className="flex items-center gap-2">
          {documentId && (
            <button
              type="button"
              onClick={handleCreateSnapshot}
              className="flex items-center gap-1.5 px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 transition-colors"
              title="Create new snapshot"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M12 4v16m8-8H4"
                />
              </svg>
              <span className="hidden sm:inline">+ Snapshot</span>
              <span className="sm:hidden">+</span>
            </button>
          )}
          <button
            type="button"
            onClick={() => setIsPanelCollapsed(true)}
            className="p-1.5 hover:bg-slate-200 rounded text-slate-500 transition-colors"
            title="Collapse panel"
          >
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M5 15l7-7 7 7-7"
              />
            </svg>
          </button>
        </div>
      </div>

      {/* Tabs */}
      <div className="flex border-b border-slate-200 bg-white overflow-x-auto">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            onClick={() => {
              setActiveTab(tab.id);
              setSearchQuery('');
            }}
            className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === tab.id
                ? 'border-blue-500 text-blue-600 bg-blue-50/50'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
            }`}
          >
            <span>{tab.icon}</span>
            <span className="hidden sm:inline">{tab.label}</span>
            <span
              className={`text-xs px-2 py-0.5 rounded-full ${
                tab.count > 0
                  ? 'bg-slate-200 text-slate-700'
                  : 'bg-slate-100 text-slate-400'
              }`}
            >
              {tab.count}
            </span>
          </button>
        ))}
      </div>

      {/* Search Bar */}
      <div className="px-4 py-3 border-b border-slate-200 bg-slate-50">
        <div className="relative">
          <svg
            className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400"
            fill="none"
            stroke="currentColor"
            viewBox="0 0 24 24"
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              strokeWidth={2}
              d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"
            />
          </svg>
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search entries..."
            className="w-full pl-10 pr-4 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          />
          {searchQuery && (
            <button
              type="button"
              onClick={() => setSearchQuery('')}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M6 18L18 6M6 6l12 12"
                />
              </svg>
            </button>
          )}
        </div>
      </div>

      {/* Content */}
      <div className="p-4 max-h-[600px] overflow-y-auto">
        {activeTab === 'history' && (
          <HistoryTab history={filteredHistory} searchQuery={searchQuery} onReferenceClick={onReferenceClick} />
        )}
        {activeTab === 'snapshots' && (
          <SnapshotsTab
            snapshots={snapshots}
            selectedSnapshotId={selectedSnapshotId}
            compareSnapshotId={compareSnapshotId}
            snapshotDiff={snapshotDiff}
            onCreateSnapshot={handleCreateSnapshot}
            onRestoreSnapshot={handleRestoreSnapshot}
            onDeleteSnapshot={handleDeleteSnapshot}
            onCompareSnapshots={handleCompareSnapshots}
          />
        )}
        {activeTab === 'dependencies' && (
          <DependenciesTab dependencies={dependencies} onReferenceClick={onReferenceClick} />
        )}
        {/* Audit Log - Simplified for now */}
        {activeTab === 'audit' && (
          <div className="text-center py-12 text-slate-500">
            <p>Audit log functionality is being updated.</p>
          </div>
        )}
      </div>
    </div>
  );
}

// History Tab Component
interface HistoryTabProps {
  history: any[];
  searchQuery: string;
  onReferenceClick?: (referenceId: string) => void;
}

function HistoryTab({ history, searchQuery, onReferenceClick }: HistoryTabProps) {
  if (history.length === 0) {
    return (
      <div className="text-center py-12">
        <div className="text-4xl mb-3">📜</div>
        <p className="text-slate-500">
          {searchQuery
            ? 'No history entries match your search'
            : 'No history available'}
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-sm font-medium text-slate-700">Change History</h3>
        <span className="text-xs text-slate-500">
          {history.length} {history.length === 1 ? 'entry' : 'entries'}
        </span>
      </div>
      {history.map((entry, index) => (
        <div
          key={index}
          className="p-4 bg-white rounded-lg border border-slate-200 hover:border-blue-300 hover:shadow-sm transition-all"
        >
          <div className="flex justify-between items-start gap-4">
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 mb-2">
                <span className="text-xs text-slate-500">
                  {new Date(entry.timestamp).toLocaleString()}
                </span>
                <span className="px-2 py-0.5 text-xs font-medium rounded-full bg-blue-100 text-blue-700 capitalize">
                  {entry.action}
                </span>
              </div>
              <div className="text-sm text-slate-700 mb-2 font-medium">
                {entry.action === 'created'
                  ? 'Reference created'
                  : entry.action === 'updated'
                  ? 'Reference updated'
                  : entry.action === 'resolved'
                  ? 'Reference resolved'
                  : entry.action === 'broken'
                  ? 'Reference broken'
                  : entry.action}
              </div>
              {entry.userId && (
                <span className="text-xs text-slate-500">
                  by <span className="font-medium text-slate-700">{entry.userId}</span>
                </span>
              )}
            </div>
          </div>
          {entry.value !== undefined && entry.oldValue !== undefined && (
            <div className="mt-3 p-3 bg-slate-50 rounded-lg">
              <div className="text-sm text-slate-600 flex items-center gap-3">
                <span className="line-through decoration-slate-400 text-slate-400">
                  {String(entry.oldValue)}
                </span>
                <svg className="w-4 h-4 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M13 7l5 5m0 0l-5-5m5 5H9"
                  />
                </svg>
                <span
                  className={`font-semibold ${
                    entry.oldValue < entry.value ? 'text-green-600' : 'text-red-600'
                  }`}
                >
                  {String(entry.value)}
                </span>
              </div>
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// Snapshots Tab Component
interface SnapshotsTabProps {
  snapshots: DocumentSnapshot[];
  selectedSnapshotId: string | null;
  compareSnapshotId: string | null;
  snapshotDiff: Diff | null;
  onCreateSnapshot: () => Promise<void>;
  onRestoreSnapshot: (id: string) => void;
  onDeleteSnapshot: (id: string) => void;
  onCompareSnapshots: (id: string) => void;
}

function SnapshotsTab({
  snapshots,
  selectedSnapshotId,
  compareSnapshotId,
  snapshotDiff,
  onCreateSnapshot,
  onRestoreSnapshot,
  onDeleteSnapshot,
  onCompareSnapshots,
}: SnapshotsTabProps) {
  if (snapshots.length === 0) {
    return (
      <div className="text-center py-12">
        <div className="text-4xl mb-3">📸</div>
        <p className="text-slate-500">No snapshots available</p>
        <p className="text-sm text-slate-400 mt-1">
          Create a snapshot to save the current document state
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-sm font-medium text-slate-700">Version Snapshots</h3>
        <span className="text-xs text-slate-500">
          {snapshots.length} {snapshots.length === 1 ? 'snapshot' : 'snapshots'}
        </span>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        {/* Snapshots List */}
        <div className="space-y-2">
          {snapshots.map((snapshot) => (
            <div
              key={snapshot.id}
              className={`p-4 rounded-lg border cursor-pointer transition-all ${
                selectedSnapshotId === snapshot.id
                  ? 'border-blue-500 bg-blue-50 ring-2 ring-blue-200'
                  : compareSnapshotId === snapshot.id
                  ? 'border-orange-500 bg-orange-50 ring-2 ring-orange-200'
                  : 'border-slate-200 hover:border-blue-300 hover:shadow-sm'
              }`}
              onClick={() => onCompareSnapshots(snapshot.id)}
            >
              <div className="flex justify-between items-start gap-3 mb-2">
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-medium text-slate-900">
                      {new Date(snapshot.timestamp).toLocaleString()}
                    </span>
                    {selectedSnapshotId === snapshot.id && (
                      <span className="text-xs px-2 py-0.5 bg-blue-600 text-white rounded">
                        Selected
                      </span>
                    )}
                  </div>
                  <div className="text-xs text-slate-500 mt-1 font-mono">
                    {snapshot.checksum.substring(0, 8)}
                  </div>
                </div>
                <div className="flex gap-1 shrink-0">
                  <button
                    type="button"
                    onClick={() => onRestoreSnapshot(snapshot.id)}
                    className="px-3 py-1.5 text-xs bg-blue-600 text-white rounded hover:bg-blue-700 transition-colors"
                    title="Restore this snapshot"
                  >
                    Restore
                  </button>
                  <button
                    type="button"
                    onClick={() => onDeleteSnapshot(snapshot.id)}
                    className="px-3 py-1.5 text-xs bg-red-50 text-red-600 hover:bg-red-100 rounded transition-colors"
                    title="Delete this snapshot"
                  >
                    Delete
                  </button>
                </div>
              </div>
            </div>
          ))}
        </div>

        {/* Snapshot Diff */}
        {snapshotDiff && (
          <div className="bg-slate-50 rounded-lg border border-slate-200 p-4">
            <div className="flex items-center gap-2 mb-3">
              <svg className="w-5 h-5 text-slate-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M9 17V7m0 10a2 2 0 012 0 0 12 0 2 2 0 012-0-2 2 2 0 0112-2 2 2 0 0112-2m-6 2l6 6 6-6"
                />
              </svg>
              <h4 className="text-sm font-medium text-slate-700">Snapshot Comparison</h4>
            </div>
            <div className="space-y-3">
              {snapshotDiff.added.length > 0 && (
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <div className="w-2 h-2 rounded-full bg-green-500" />
                    <div className="text-xs font-medium text-green-700">Added References</div>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {snapshotDiff.added.map((id) => (
                      <span
                        key={id}
                        className="inline-flex items-center gap-1 px-2 py-1 bg-green-100 text-green-800 text-xs rounded"
                      >
                        <span className="font-mono">{id.substring(0, 8)}</span>
                      </span>
                    ))}
                  </div>
                </div>
              )}
              {snapshotDiff.removed.length > 0 && (
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <div className="w-2 h-2 rounded-full bg-red-500" />
                    <div className="text-xs font-medium text-red-700">Removed References</div>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {snapshotDiff.removed.map((id) => (
                      <span
                        key={id}
                        className="inline-flex items-center gap-1 px-2 py-1 bg-red-100 text-red-800 text-xs rounded"
                      >
                        <span className="font-mono">{id.substring(0, 8)}</span>
                      </span>
                    ))}
                  </div>
                </div>
              )}
              {snapshotDiff.modified.length > 0 && (
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <div className="w-2 h-2 rounded-full bg-orange-500" />
                    <div className="text-xs font-medium text-orange-700">Modified References</div>
                  </div>
                  <div className="space-y-2">
                    {snapshotDiff.modified.map((change, index) => (
                      <div
                        key={index}
                        className="p-3 bg-white rounded-lg border border-slate-200"
                      >
                        <div className="text-xs text-slate-500 mb-1 font-mono">
                          {change.path}
                        </div>
                        <div className="flex items-center gap-3 text-sm">
                          <span className="line-through decoration-slate-400 text-slate-400">
                            {String(change.oldValue)}
                          </span>
                          <svg className="w-4 h-4 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth={2}
                              d="M13 7l5 5m0 0l-5-5m5 5H9"
                            />
                          </svg>
                          <span className="font-medium text-slate-900">
                            {String(change.newValue)}
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// Dependencies Tab Component
interface DependenciesTabProps {
  dependencies: string[];
  onReferenceClick?: (referenceId: string) => void;
}

function DependenciesTab({ dependencies, onReferenceClick }: DependenciesTabProps) {
  if (dependencies.length === 0) {
    return (
      <div className="text-center py-12">
        <div className="text-4xl mb-3">🔗</div>
        <p className="text-slate-500">No dependencies found</p>
        <p className="text-sm text-slate-400 mt-1">
          This reference is not used by any other documents or cells
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h3 className="text-sm font-medium text-slate-700">Reference Dependencies</h3>
          <p className="text-sm text-slate-500 mt-1">
            Documents and references that depend on this data
          </p>
        </div>
        <span className="text-xs text-slate-500">
          {dependencies.length} {dependencies.length === 1 ? 'dependency' : 'dependencies'}
        </span>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
        {dependencies.map((depId, index) => (
          <button
            key={depId}
            type="button"
            onClick={() => onReferenceClick?.(depId)}
            className="text-left p-4 bg-white rounded-lg border border-slate-200 hover:border-blue-300 hover:shadow-sm transition-all"
          >
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 flex items-center justify-center rounded-lg bg-slate-100">
                <svg className="w-5 h-5 text-slate-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 01.489-2.03l4-4a4 4 0 10-5.656 0l-4 4a4 4 0 01-.489 2.03m.818-2.982a4 4 0 01-.488-2.029l-4-4a4 4 0 10-5.656 0l-4 4a4 4 0 10-.489 2.03m-4-4l4 4a4 4 0 01.489 2.029"
                  />
                </svg>
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-xs font-mono text-slate-500 truncate">
                  {depId.substring(0, 12)}{depId.length > 12 ? '...' : ''}
                </div>
                <div className="text-xs text-slate-600 mt-0.5">Reference ID</div>
              </div>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}
