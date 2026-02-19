/**
 * Data Update Confirmation Component
 *
 * Shows confirmation dialogs for batch updates
 */

'use client';

import React, { useState, useEffect } from 'react';
import type { Conflict, ResolveStrategy } from '../../core/sync';

interface DataUpdateConfirmationProps {
  conflicts: Conflict[];
  onResolve: (conflictId: string, strategy: ResolveStrategy) => void;
  onResolveAll: (strategy: ResolveStrategy) => void;
  onCancel: () => void;
  onClose?: () => void;
}

interface ConflictItemProps {
  conflict: Conflict;
  onResolve: (conflictId: string, strategy: ResolveStrategy) => void;
}

function ConflictItem({ conflict, onResolve }: ConflictItemProps) {
  const [selectedStrategy, setSelectedStrategy] = useState<ResolveStrategy | null>(null);

  const handleResolve = (strategy: ResolveStrategy) => {
    setSelectedStrategy(strategy);
    onResolve(conflict.id, strategy);
  };

  const formatDate = (timestamp: number) => {
    return new Date(timestamp).toLocaleString();
  };

  const renderValue = (value: any) => {
    if (value === null || value === undefined) return 'N/A';
    if (typeof value === 'object') return JSON.stringify(value).slice(0, 50) + '...';
    return String(value);
  };

  const getTypeIcon = (type: string) => {
    switch (type) {
      case 'concurrent-edit':
        return '⚡';
      case 'version-mismatch':
        return '📅';
      case 'data-inconsistency':
        return '⚠️';
      default:
        return '❓';
    }
  };

  return (
    <div className="border border-gray-200 rounded-lg p-4 mb-3 bg-white">
      <div className="flex items-start gap-3">
        <div className="text-2xl">{getTypeIcon(conflict.type)}</div>
        <div className="flex-1">
          <div className="flex items-center justify-between mb-2">
            <h3 className="font-medium text-gray-900">
              Conflict: {conflict.referenceId}
            </h3>
            <span className={`text-xs px-2 py-1 rounded ${
              conflict.type === 'data-inconsistency'
                ? 'bg-red-100 text-red-700'
                : 'bg-yellow-100 text-yellow-700'
            }`}>
              {conflict.type.replace('-', ' ').toUpperCase()}
            </span>
          </div>

          <div className="grid grid-cols-2 gap-4 text-sm">
            <div>
              <div className="text-gray-500 text-xs mb-1">Local Value</div>
              <div className="font-mono bg-gray-50 p-2 rounded">
                {renderValue(conflict.localValue)}
              </div>
              <div className="text-gray-400 text-xs mt-1">
                Updated: {formatDate(conflict.localTimestamp)}
              </div>
            </div>
            <div>
              <div className="text-gray-500 text-xs mb-1">Remote Value</div>
              <div className="font-mono bg-blue-50 p-2 rounded">
                {renderValue(conflict.remoteValue)}
              </div>
              <div className="text-gray-400 text-xs mt-1">
                Updated: {formatDate(conflict.remoteTimestamp)}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="mt-3 pt-3 border-t border-gray-100">
        <div className="text-gray-500 text-xs mb-2">Resolution:</div>
        <div className="flex gap-2 flex-wrap">
          <button
            onClick={() => handleResolve('keep-local')}
            className="px-3 py-1.5 text-sm rounded border border-gray-300 hover:bg-gray-50 transition-colors"
          >
            Keep Local
          </button>
          <button
            onClick={() => handleResolve('keep-remote')}
            className="px-3 py-1.5 text-sm rounded border border-blue-300 bg-blue-50 text-blue-700 hover:bg-blue-100 transition-colors"
          >
            Use Remote
          </button>
          <button
            onClick={() => handleResolve('merge')}
            className="px-3 py-1.5 text-sm rounded border border-green-300 bg-green-50 text-green-700 hover:bg-green-100 transition-colors"
          >
            Merge
          </button>
          {conflict.suggestedStrategy && (
            <button
              onClick={() => handleResolve(conflict.suggestedStrategy!)}
              className="px-3 py-1.5 text-sm rounded border border-purple-300 bg-purple-50 text-purple-700 hover:bg-purple-100 transition-colors"
            >
              Suggested
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

export function DataUpdateConfirmation({
  conflicts,
  onResolve,
  onResolveAll,
  onCancel,
  onClose,
}: DataUpdateConfirmationProps) {
  const [showDetails, setShowDetails] = useState(false);

  const handleResolveAll = (strategy: ResolveStrategy) => {
    onResolveAll(strategy);
    if (onClose) onClose();
  };

  const getStrategyLabel = (strategy: ResolveStrategy) => {
    switch (strategy) {
      case 'keep-local':
        return 'Keep your local changes';
      case 'keep-remote':
        return 'Use the remote data';
      case 'merge':
        return 'Merge both versions';
      case 'latest-timestamp':
        return 'Use the most recently updated';
      case 'manual':
        return 'Manual resolution required';
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
      <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full max-h-[90vh] overflow-hidden flex flex-col">
        {/* Header */}
        <div className="p-6 border-b border-gray-100">
          <div className="flex items-center justify-between">
            <div>
              <h2 className="text-xl font-semibold text-gray-900">
                Data Update Conflicts
              </h2>
              <p className="text-gray-500 text-sm mt-1">
                {conflicts.length} conflict{conflicts.length !== 1 ? 's' : ''} need resolution
              </p>
            </div>
            <button
              onClick={onClose}
              className="text-gray-400 hover:text-gray-600 transition-colors"
            >
              <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          </div>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-6">
          <div className="mb-4 bg-blue-50 border border-blue-100 rounded-lg p-4">
            <div className="flex items-start gap-3">
              <svg className="w-5 h-5 text-blue-500 flex-shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <div className="text-sm text-blue-800">
                <p className="font-medium mb-1">Data updates are pending</p>
                <p className="text-blue-600">
                  Choose how to resolve each conflict or apply a resolution strategy to all.
                </p>
              </div>
            </div>
          </div>

          <div className="mb-4">
            <button
              onClick={() => setShowDetails(!showDetails)}
              className="text-sm text-gray-600 hover:text-gray-800 flex items-center gap-1"
            >
              <svg className={`w-4 h-4 transition-transform ${showDetails ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
              </svg>
              {showDetails ? 'Hide Details' : 'Show Details'}
            </button>
          </div>

          {showDetails ? (
            <div className="space-y-3">
              {conflicts.map(conflict => (
                <ConflictItem
                  key={conflict.id}
                  conflict={conflict}
                  onResolve={onResolve}
                />
              ))}
            </div>
          ) : (
            <div className="text-center text-gray-500 py-8">
              <svg className="w-16 h-16 mx-auto mb-3 text-gray-300" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h10a2 2 0 012 2v14a2 2 0 01-2 2h-2" />
              </svg>
              <p>Details hidden. Click "Show Details" to view conflicts.</p>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-6 border-t border-gray-100 bg-gray-50">
          <div className="flex items-center justify-between">
            <button
              onClick={onCancel}
              className="px-4 py-2 text-sm font-medium text-gray-700 hover:text-gray-900 transition-colors"
            >
              Cancel
            </button>
            <div className="flex gap-2">
              {conflicts.length > 1 && (
                <>
                  <button
                    onClick={() => handleResolveAll('keep-local')}
                    className="px-4 py-2 text-sm font-medium rounded-lg border border-gray-300 bg-white hover:bg-gray-50 transition-colors"
                  >
                    Keep All Local
                  </button>
                  <button
                    onClick={() => handleResolveAll('keep-remote')}
                    className="px-4 py-2 text-sm font-medium rounded-lg border border-blue-300 bg-blue-50 text-blue-700 hover:bg-blue-100 transition-colors"
                  >
                    Use All Remote
                  </button>
                  <button
                    onClick={() => handleResolveAll('merge')}
                    className="px-4 py-2 text-sm font-medium rounded-lg border border-green-300 bg-green-50 text-green-700 hover:bg-green-100 transition-colors"
                  >
                    Merge All
                  </button>
                </>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
