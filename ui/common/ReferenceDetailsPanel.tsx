'use client';

import { useState, useEffect } from 'react';
import type { DataReference } from '@/core/types';
import { DataLinkManager } from '@/core/editors/DataLinkManager';
import { ReferenceState } from '@/core/types';
import { eventBus, Events } from '@/core/eventBus';

interface ReferenceDetailsPanelProps {
  referenceId: string | null;
  dataLinkManager: DataLinkManager | null;
  onClose: () => void;
}

/**
 * Reference Details Panel Component
 *
 * Shows detailed information about a selected data reference,
 * including refresh and delete functionality.
 */
export function ReferenceDetailsPanel({
  referenceId,
  dataLinkManager,
  onClose,
}: ReferenceDetailsPanelProps) {
  const [reference, setReference] = useState<DataReference | null>(null);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [refreshError, setRefreshError] = useState<string | null>(null);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);

  // Load reference details when referenceId changes
  useEffect(() => {
    if (referenceId && dataLinkManager) {
      const ref = dataLinkManager.getReference(referenceId);
      setReference(ref || null);
      setRefreshError(null);
    } else {
      setReference(null);
    }
  }, [referenceId, dataLinkManager]);

  // Listen for reference updates
  useEffect(() => {
    if (!referenceId) return;

    const handleReferenceUpdate = (data: any) => {
      if (data.referenceId === referenceId && dataLinkManager) {
        const ref = dataLinkManager.getReference(referenceId);
        setReference(ref || null);
        setRefreshError(null);
      }
    };

    eventBus.on(Events.ReferenceUpdated as any, handleReferenceUpdate);

    return () => {
      eventBus.off(Events.ReferenceUpdated as any, handleReferenceUpdate);
    };
  }, [referenceId, dataLinkManager]);

  const handleRefresh = async () => {
    if (!referenceId || !dataLinkManager || isRefreshing) return;

    setIsRefreshing(true);
    setRefreshError(null);

    try {
      // Emit refresh event - will be handled by DataSyncEngine
      eventBus.emit(Events.ReferenceUpdated as any, {
        referenceId,
        action: 'refresh',
      });

      // Trigger actual refresh via data link manager
      const ref = dataLinkManager.getReference(referenceId);
      if (ref) {
        await dataLinkManager.resolveReference(ref);
        // Reload reference after refresh
        const updatedRef = dataLinkManager.getReference(referenceId);
        setReference(updatedRef || null);
      }
    } catch (error) {
      console.error('Error refreshing reference:', error);
      setRefreshError(error instanceof Error ? error.message : 'Failed to refresh reference');
    } finally {
      setIsRefreshing(false);
    }
  };

  const handleDelete = async () => {
    if (!referenceId || !dataLinkManager) return;

    setIsDeleting(true);
    try {
      dataLinkManager.deleteReference(referenceId);
      setShowDeleteConfirm(false);
      onClose();
    } catch (error) {
      console.error('Error deleting reference:', error);
      setIsDeleting(false);
      // Keep dialog open on error
      setRefreshError(error instanceof Error ? error.message : 'Failed to delete reference');
    }
  };

  const navigateToSource = () => {
    if (!reference) return;

    // Emit event for navigation to Excel cell
    eventBus.emit(Events.DataChanged as any, {
      action: 'navigate',
      target: reference.source,
    });

    // TODO: Implement actual navigation to Excel cell
    console.log('Navigate to:', reference.source);
  };

  const getStateStyle = (state: ReferenceState) => {
    switch (state) {
      case 'active':
        return { bg: 'bg-green-50', border: 'border-green-200', text: 'text-green-700', icon: '✓', dot: 'bg-green-500' };
      case 'stale':
        return { bg: 'bg-yellow-50', border: 'border-yellow-200', text: 'text-yellow-700', icon: '⏳', dot: 'bg-yellow-500' };
      case 'broken':
        return { bg: 'bg-red-50', border: 'border-red-200', text: 'text-red-700', icon: '✗', dot: 'bg-red-500' };
      case 'conflict':
        return { bg: 'bg-orange-50', border: 'border-orange-200', text: 'text-orange-700', icon: '⚠', dot: 'bg-orange-500' };
      default:
        return { bg: 'bg-slate-50', border: 'border-slate-200', text: 'text-slate-700', icon: '?', dot: 'bg-slate-500' };
    }
  };

  if (!reference) {
    return (
      <div className="bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-slate-50 border-b border-slate-200">
          <div className="flex items-center gap-2">
            <span className="text-xl">📋</span>
            <h2 className="text-lg font-semibold text-slate-900">Reference Details</h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-400 hover:text-slate-600 transition-colors"
            title="Close panel"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
        {/* Content */}
        <div className="p-8 text-center">
          <div className="text-4xl mb-4">🔍</div>
          <p className="text-slate-500">Select a reference to view details</p>
          <p className="text-sm text-slate-400 mt-2">
            Click on any reference in your document to see its details
          </p>
        </div>
      </div>
    );
  }

  const stateStyle = getStateStyle(reference.state);

  return (
    <div className="bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
      {/* Header */}
      <div className="flex items-center justify-between px-4 py-3 bg-slate-50 border-b border-slate-200">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-full flex items-center justify-center" style={{ backgroundColor: stateStyle.bg.replace('bg-', 'bg-').replace('-50', '-100') }}>
            <span className="text-xl">{stateStyle.icon}</span>
          </div>
          <div>
            <h2 className="text-lg font-semibold text-slate-900">Reference Details</h2>
            <p className="text-sm text-slate-500 mt-0.5">
              {reference.state === 'active' && 'Reference is up to date'}
              {reference.state === 'stale' && 'Reference may need updating'}
              {reference.state === 'broken' && 'Reference is broken'}
              {reference.state === 'conflict' && 'Reference has conflicts'}
            </p>
          </div>
        </div>
        <button
          type="button"
          onClick={onClose}
          className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-400 hover:text-slate-600 transition-colors"
          title="Close panel"
        >
          <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
          </svg>
        </button>
      </div>

      {/* Content */}
      <div className="p-4">
        {/* State Badge */}
        <div className={`inline-flex items-center gap-2 px-4 py-2 rounded-full border mb-6 ${stateStyle.bg} ${stateStyle.border}`}>
          <div className="w-2 h-2 rounded-full" style={{ backgroundColor: stateStyle.dot }}></div>
          <span className="font-medium capitalize text-base">{reference.state}</span>
        </div>

        {/* Value Display */}
        <div className="mb-6">
          <label className="block text-sm font-medium text-slate-700 mb-2">Current Value</label>
          <div className={`p-4 rounded-lg border transition-all ${
            isRefreshing
              ? 'border-blue-400 bg-blue-50/50'
              : 'border-slate-200 bg-slate-50'
          }`}>
            {isRefreshing ? (
              <div className="flex items-center justify-center gap-3">
                <svg className="animate-spin w-5 h-5 text-slate-400" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path
                    className="opacity-75"
                    fill="currentColor"
                    d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                  />
                </svg>
                <span className="text-slate-500">Refreshing...</span>
              </div>
            ) : (
              <span className="text-2xl font-semibold text-slate-900">
                {reference.display.value !== undefined ? String(reference.display.value) : 'N/A'}
              </span>
            )}
          </div>
        </div>

        {/* Expression */}
        {reference.display.expression && (
          <div className="mb-6">
            <label className="block text-sm font-medium text-slate-700 mb-2">Expression</label>
            <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
              <code className="text-sm text-slate-700">{reference.display.expression}</code>
            </div>
          </div>
        )}

        {/* Source Information */}
        <div className="mb-6">
          <label className="block text-sm font-medium text-slate-700 mb-2">Source Location</label>
          <div className="p-4 bg-blue-50 rounded-lg border border-blue-200">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-sm">
              <div className="flex items-center gap-2">
                <svg className="w-4 h-4 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 0v8m6-8h6m-6 0v-8m1.5 1.5 21-12" />
                </svg>
                <span className="text-slate-500">File:</span>
                <span className="ml-2 font-medium text-slate-900 truncate">{reference.source.fileName}</span>
              </div>
              <div className="flex items-center gap-2">
                <svg className="w-4 h-4 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 11H5m14 0a2 2 0 012-2h10a2 2 0 012-2V5a2 2 0 00-2-2h5a2 2 0 010-2v10a2 2 0 012 2h5a2 2 0 012-2z" />
                </svg>
                <span className="text-slate-500">Sheet:</span>
                <span className="ml-2 font-medium text-slate-900">{reference.source.sheetName}</span>
              </div>
              <div className="flex items-center gap-2">
                <svg className="w-4 h-4 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 21a4 4 0 11-4 4 4 0 0118 0z" />
                </svg>
                <span className="text-slate-500">Type:</span>
                <span className="ml-2 font-medium text-slate-900 capitalize">{reference.type}</span>
              </div>
              <div className="flex items-center gap-2">
                <svg className="w-4 h-4 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 7h6m0 2v8m-6-2h6m2 0v8M7 10l5 5m0 0l-5-5m5 5H9" />
                </svg>
                <span className="text-slate-500">Formula:</span>
                <span className="ml-2 font-medium text-slate-900">
                  {reference.source.isFormula ? (
                    <span className="text-blue-600">Yes</span>
                  ) : (
                    <span className="text-slate-500">No</span>
                  )}
                </span>
              </div>
            </div>
            <button
              type="button"
              onClick={navigateToSource}
              className="mt-4 w-full sm:w-auto flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-medium bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-all hover:shadow-md active:scale-[0.98]"
              disabled={isRefreshing}
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" />
              </svg>
              <span className="hidden sm:inline">Navigate to Source</span>
              <span className="sm:hidden">Go to Source</span>
            </button>
          </div>
        </div>

        {/* Reference ID */}
        <div className="mb-6">
          <label className="block text-sm font-medium text-slate-700 mb-2">Reference ID</label>
          <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
            <div className="flex items-center gap-2">
              <code className="text-sm text-slate-600 font-mono flex-1">{reference.id}</code>
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(reference.id);
                }}
                className="p-1.5 rounded hover:bg-slate-200 text-slate-400 hover:text-slate-600 transition-colors"
                title="Copy to clipboard"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 16H6a2 2 0 01-2-2V4a2 2 0 012 2h5a2 2 0 010-2v10a2 2 0 012 2h5a2 2 0 012 2v5a2 2 0 01-2 2h5a2 2 0 010-2v10a2 2 0 012 2z" />
                </svg>
              </button>
            </div>
          </div>
        </div>

        {/* Metadata */}
        {reference.metadata && (
          <div className="mb-6">
            <details className="group" open>
              <summary className="flex items-center gap-2 cursor-pointer hover:text-slate-900 text-sm font-medium text-slate-700 list-none">
                <svg className="w-4 h-4 group-open:rotate-180 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                </svg>
                Metadata
              </summary>
              <div className="p-3 mt-2 bg-slate-50 rounded-lg border border-slate-200 space-y-2 text-sm">
                {reference.metadata.created && (
                  <div className="flex items-center gap-3">
                    <svg className="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m-6-6l6 6m6 0V8h-4z" />
                    </svg>
                    <div>
                      <span className="text-slate-500 w-24">Created:</span>
                      <span className="text-slate-700 font-medium">{new Date(reference.metadata.created).toLocaleString()}</span>
                    </div>
                  </div>
                )}
                {reference.metadata.updated && (
                  <div className="flex items-center gap-3">
                    <svg className="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v16a2 2 0 012-2h16a2 2 0 012-2V4a2 2 0 00-2-2h5a2 2 0 010-2v10a2 2 0 012 2h5a2 2 0 012 2z" />
                    </svg>
                    <div>
                      <span className="text-slate-500 w-24">Updated:</span>
                      <span className="text-slate-700 font-medium">{new Date(reference.metadata.updated).toLocaleString()}</span>
                    </div>
                  </div>
                )}
                {reference.metadata.author && (
                  <div className="flex items-center gap-3">
                    <svg className="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M20 21v-2a4 4 0 00-8 4H4a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2V5a2 2 0 01-2 2h5a2 2 0 010-2v10a2 2 0 012 2z" />
                    </svg>
                    <div>
                      <span className="text-slate-500 w-24">Author:</span>
                      <span className="text-slate-700 font-medium">{reference.metadata.author}</span>
                    </div>
                  </div>
                )}
              </div>
            </details>
          </div>
        )}

        {/* Error Message */}
        {refreshError && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg animate-fade-in">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center shrink-0">
                <svg className="w-6 h-6 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <div className="flex-1">
                <p className="text-sm font-medium text-red-900">Error</p>
                <p className="text-sm text-red-700 mt-1">{refreshError}</p>
              </div>
            </div>
          </div>
        )}

        {/* Action Buttons */}
        <div className="flex gap-3">
          <button
            type="button"
            onClick={handleRefresh}
            disabled={isRefreshing}
            className={`flex-1 flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-medium rounded-lg transition-all ${
              isRefreshing
                ? 'bg-slate-200 text-slate-500 cursor-not-allowed'
                : 'bg-blue-600 text-white hover:bg-blue-700 hover:shadow-md active:scale-[0.98]'
            }`}
          >
            {isRefreshing ? (
              <>
                <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                  <path
                    className="opacity-75"
                    fill="currentColor"
                    d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                  />
                </svg>
                <span>Refreshing...</span>
              </>
            ) : (
              <>
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                </svg>
                <span className="hidden sm:inline">Refresh Value</span>
                <span className="sm:hidden">Refresh</span>
              </>
            )}
          </button>

          <button
            type="button"
            onClick={() => setShowDeleteConfirm(true)}
            className="px-4 py-2.5 text-sm font-medium text-red-600 bg-red-50 hover:bg-red-100 rounded-lg transition-all hover:shadow-md active:scale-[0.98] border border-red-200"
            disabled={isRefreshing}
          >
            Delete
          </button>
        </div>

        {/* Delete Confirmation Dialog */}
        {showDeleteConfirm && (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[100] animate-scale-in">
            <div className="bg-white rounded-xl shadow-2xl border border-red-200 p-6 max-w-sm mx-4">
              <div className="flex items-start gap-3 mb-4">
                <div className="w-12 h-12 rounded-full bg-red-100 flex items-center justify-center shrink-0">
                  <svg className="w-6 h-6 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                </div>
                <div>
                  <h3 className="text-base font-semibold text-slate-900">Delete Reference?</h3>
                  <p className="text-sm text-slate-600 mt-1">
                    This action cannot be undone. The reference will be removed from the document.
                  </p>
                </div>
              </div>
              <div className="flex gap-3">
                <button
                  type="button"
                  onClick={() => setShowDeleteConfirm(false)}
                  className="flex-1 px-4 py-2.5 text-sm font-medium bg-slate-100 text-slate-700 rounded-lg hover:bg-slate-200 transition-colors border border-slate-300"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleDelete}
                  disabled={isDeleting}
                  className={`flex-1 px-4 py-2.5 text-sm font-medium bg-red-600 text-white rounded-lg transition-colors ${
                    isDeleting ? 'opacity-70 cursor-not-allowed' : 'hover:bg-red-700'
                  }`}
                >
                  {isDeleting ? (
                    <>
                      <svg className="animate-spin w-4 h-4 inline" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                        />
                      </svg>
                      Deleting...
                    </>
                  ) : (
                    'Delete'
                  )}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
