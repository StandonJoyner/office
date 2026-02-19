'use client';

import { useEffect, useRef, useState } from 'react';
import { eventBus, Events } from '@/core/eventBus';
import { DataChange } from '@/core/types';

interface DataUpdateNotificationProps {
  /**
   * Callback when a data change is detected
   */
  onDataChange?: (changes: DataChange[]) => void;
}

/**
 * Data Update Notification Component
 *
 * Displays a toast notification when data is updated via DataSyncEngine.
 */
export function DataUpdateNotification({ onDataChange }: DataUpdateNotificationProps) {
  const [showNotification, setShowNotification] = useState(false);
  const [changes, setChanges] = useState<DataChange[]>([]);
  const [isHiding, setIsHiding] = useState(false);
  const timeoutIdRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const onDataChangeRef = useRef(onDataChange);
  onDataChangeRef.current = onDataChange;

  useEffect(() => {
    const handleDataChanged = (data: DataChange[]) => {
      setChanges(data);
      setShowNotification(true);
      setIsHiding(false);
      onDataChangeRef.current?.(data);

      // Clear previous auto-hide timer
      if (timeoutIdRef.current) {
        clearTimeout(timeoutIdRef.current);
        timeoutIdRef.current = null;
      }

      // Auto-hide after 5 seconds
      timeoutIdRef.current = setTimeout(() => {
        setIsHiding(true);
        const hideId = setTimeout(() => {
          setShowNotification(false);
        }, 300);
        timeoutIdRef.current = hideId;
      }, 5000);
    };

    eventBus.on(Events.DataChanged, handleDataChanged);

    return () => {
      eventBus.off(Events.DataChanged, handleDataChanged);
      if (timeoutIdRef.current) {
        clearTimeout(timeoutIdRef.current);
        timeoutIdRef.current = null;
      }
    };
  }, []);

  const handleDismiss = () => {
    setIsHiding(true);
    setTimeout(() => {
      setShowNotification(false);
    }, 300);
  };

  const handleRefreshAll = () => {
    onDataChange?.(changes);
  };

  if (!showNotification || changes.length === 0) {
    return null;
  }

  return (
    <div
      className={`fixed bottom-4 right-4 z-50 ${
        isHiding ? 'animate-slide-out' : 'animate-slide-in'
      }`}
    >
      <div className="bg-white rounded-xl shadow-2xl border border-slate-200 overflow-hidden max-w-md min-w-[320px]">
        {/* Header */}
        <div className="flex items-center justify-between px-4 py-3 bg-gradient-to-r from-slate-50 to-slate-100 border-b border-slate-200">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-green-100 flex items-center justify-center">
              <svg
                className="w-6 h-6 text-green-600"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M5 13l4 4L19 7"
                />
              </svg>
            </div>
            <div>
              <h3 className="text-base font-semibold text-slate-900">Data Updated</h3>
              <p className="text-sm text-slate-500">
                {changes.length} reference{changes.length > 1 ? 's' : ''} refreshed
              </p>
            </div>
          </div>
          <button
            onClick={handleDismiss}
            className="p-1.5 rounded-lg hover:bg-slate-200 text-slate-400 hover:text-slate-600 transition-colors"
            aria-label="Dismiss notification"
            title="Dismiss"
          >
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        {/* Content */}
        <div className="max-h-[300px] overflow-y-auto p-4">
          {/* Show up to 3 changes with full details */}
          {changes.slice(0, 3).map((change, index) => (
            <div
              key={change.referenceId}
              className={`mb-3 last:mb-0 p-3 bg-slate-50 rounded-lg border border-slate-200 ${
                index === 0 ? 'animate-scale-in' : ''
              }`}
            >
              <div className="flex items-start justify-between gap-3">
                <div className="flex-1 min-w-0">
                  <div className="text-xs text-slate-500 font-mono mb-1 truncate">
                    {change.referenceId}
                  </div>
                  <div className="flex items-center gap-2 text-sm">
                    <span className="line-through decoration-red-300 text-slate-400 bg-red-50 px-1.5 py-0.5 rounded">
                      {String(change.oldValue)}
                    </span>
                    <svg className="w-4 h-4 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7l5 5m0 0l-5-5m5 5H9" />
                    </svg>
                    <span className="font-semibold text-green-600 bg-green-50 px-1.5 py-0.5 rounded">
                      {String(change.newValue)}
                    </span>
                  </div>
                  <div className="text-xs text-slate-500 mt-1">
                    {new Date(change.timestamp).toLocaleTimeString()}
                  </div>
                </div>
              </div>
            </div>
          ))}

          {/* Show "and X more" if there are more changes */}
          {changes.length > 3 && (
            <div className="pt-2 border-t border-slate-200">
              <button
                onClick={handleRefreshAll}
                className="w-full flex items-center justify-center gap-2 px-4 py-2 text-sm bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
              >
                <svg className="w-4 h-4 text-slate-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v16a2 2 0 012-2h16a2 2 0 012-2V4a2 2 0 00-2-2h5a2 2 0 010-2v10a2 2 0 012 2h5a2 2 0 012 2v5a2 2 0 00-2 2z" />
                </svg>
                View all {changes.length - 3} more change{changes.length - 3 > 1 ? 's' : ''}
              </button>
            </div>
          )}

          {/* Footer action */}
          <div className="pt-3 mt-4 border-t border-slate-200">
            <button
              onClick={handleDismiss}
              className="w-full flex items-center justify-center gap-2 px-4 py-2.5 text-sm font-medium bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
              Dismiss
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
