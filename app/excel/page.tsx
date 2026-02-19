'use client';

/**
 * Excel Editor Integration Demo
 *
 * Demonstrates integration between ExcelEditor and DataLinkManager.
 */

import React, { useRef, useState, useEffect, useCallback } from 'react';
import { ExcelEditor, type ExcelEditorRefType } from '@/ui/excel';
import {
  createDataLinkManager,
  createUniverWorkbookAdapter,
  type DataLinkManager,
  type DataReference,
  type DataSource,
} from '../../core';

export default function ExcelPage() {
  console.log('[ExcelPage] RENDER');

  const excelEditorRef = useRef<ExcelEditorRefType>(null);
  const dataLinkManagerRef = useRef<DataLinkManager>(createDataLinkManager());
  const [references, setReferences] = useState<DataReference[]>([]);
  const [selectedCell, setSelectedCell] = useState<{ sheetId: string; row: number; column: number } | null>(null);
  const [isReady, setIsReady] = useState(false);

  // Track component lifecycle
  useEffect(() => {
    console.log('[ExcelPage] MOUNT');
    return () => {
      console.log('[ExcelPage] UNMOUNT');
    };
  }, []);

  /**
   * Refresh all references
   */
  const refreshReferences = useCallback(async () => {
    const refs = references;
    const promises = refs.map(ref => dataLinkManagerRef.current.resolveReference(ref));
    await Promise.all(promises);

    // Update references list with new values
    setReferences(dataLinkManagerRef.current.getAllReferences());
  }, [references]);

  useEffect(() => {
    // Set up data link manager listeners
    const manager = dataLinkManagerRef.current;

    // Listen for data changes
    const handleDataSourceChange = () => {
      console.log('Data source changed');
      // Update references when data changes
      refreshReferences();
    };

    manager.onDataSourceChange(handleDataSourceChange);

    return () => {
      // Cleanup
      manager.offDataSourceChange(handleDataSourceChange);
    };
  }, [refreshReferences]);

  /**
   * Handle workbook loaded event
   */
  const handleWorkbookLoaded = (workbook: any) => {
    console.log('Workbook loaded:', workbook);

    // Get the API from the editor ref
    const api = excelEditorRef.current?.getAPI();

    if (api && workbook) {
      // Create an adapter that bridges FUniver events to IUniverWorkbook interface
      const adapter = createUniverWorkbookAdapter(workbook as any, api as any);
      // Set workbook adapter in data link manager
      dataLinkManagerRef.current.setExcelWorkbook(adapter);
    }

    setIsReady(true);
  };

  /**
   * Handle cell selection
   */
  const handleCellSelected = (sheetId: string, row: number, column: number) => {
    setSelectedCell({ sheetId, row, column });
  };

  /**
   * Handle range selection
   */
  const handleRangeSelected = (sheetId: string, range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
    console.log('Range selected:', range);
    setSelectedCell({ sheetId, row: range.startRow, column: range.startCol });
  };

  /**
   * Create a reference to selected cell
   */
  const createReference = () => {
    if (!selectedCell) {
      alert('Please select a cell first');
      return;
    }

    const editor = excelEditorRef.current;
    if (!editor) return;

    const workbook = editor.getWorkbook();
    if (!workbook) return;

    const sheet = workbook.getActiveSheet();
    if (!sheet) return;

    const dataSource: DataSource = {
      fileId: workbook.getId(),
      fileName: workbook.getName(),
      sheetId: selectedCell.sheetId,
      sheetName: sheet.getName(),
      range: {
        startRow: selectedCell.row,
        startCol: selectedCell.column,
      },
      isFormula: false,
    };

    const refId = dataLinkManagerRef.current.createReference(dataSource, 'demo-document');

    // Update references list
    setReferences([...references, dataLinkManagerRef.current.getReference(refId)!]);

    alert(`Reference created: ${refId}`);
  };

  /**
   * Delete a reference
   */
  const deleteReference = (id: string) => {
    dataLinkManagerRef.current.deleteReference(id);
    setReferences(references.filter(ref => ref.id !== id));
  };

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-3xl font-bold text-gray-900 mb-6">Excel Editor Demo</h1>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Excel Editor */}
          <div className="lg:col-span-2">
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-xl font-semibold mb-4">Spreadsheet</h2>
              <ExcelEditor
                ref={excelEditorRef as any}
                height="600px"
                enableDataSource={true}
                onWorkbookLoaded={handleWorkbookLoaded}
                onCellSelected={handleCellSelected}
                onRangeSelected={handleRangeSelected}
              />
            </div>
          </div>

          {/* Data Source Panel */}
          <div className="lg:col-span-1">
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-xl font-semibold mb-4">Data References</h2>

              {/* Selected Cell Info */}
              {selectedCell && (
                <div className="mb-4 p-3 bg-blue-50 rounded-lg">
                  <p className="text-sm font-medium text-blue-900">Selected Cell</p>
                  <p className="text-blue-700">
                    Sheet: {selectedCell.sheetId}, Row: {selectedCell.row + 1}, Col: {selectedCell.column + 1}
                  </p>
                  <button
                    onClick={createReference}
                    className="mt-2 px-3 py-1.5 bg-blue-600 text-white text-sm rounded hover:bg-blue-700 transition-colors"
                  >
                    Create Reference
                  </button>
                </div>
              )}

              {/* References List */}
              <div className="mb-4">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="text-sm font-medium text-gray-700">Active References</h3>
                  <button
                    onClick={refreshReferences}
                    className="text-sm text-blue-600 hover:text-blue-700"
                  >
                    Refresh
                  </button>
                </div>

                {references.length === 0 ? (
                  <p className="text-sm text-gray-500">No references created yet</p>
                ) : (
                  <div className="space-y-2">
                    {references.map(ref => (
                      <div key={ref.id} className="p-3 bg-gray-50 rounded-lg">
                        <div className="flex items-center justify-between">
                          <div>
                            <p className="text-sm font-medium text-gray-900">
                              {ref.source.fileName} / {ref.source.sheetName}
                            </p>
                            <p className="text-xs text-gray-500">
                              Cell: ({ref.source.range.startRow + 1}, {ref.source.range.startCol + 1})
                            </p>
                            <p className="text-xs text-gray-500 mt-1">
                              Value: <span className="font-medium">{ref.display.value ?? 'N/A'}</span>
                            </p>
                          </div>
                          <button
                            onClick={() => deleteReference(ref.id)}
                            className="text-red-600 hover:text-red-700"
                          >
                            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                            </svg>
                          </button>
                        </div>
                        <span className={`inline-block px-2 py-0.5 text-xs rounded mt-2 ${
                          ref.state === 'active' ? 'bg-green-100 text-green-800' :
                          ref.state === 'stale' ? 'bg-yellow-100 text-yellow-800' :
                          'bg-red-100 text-red-800'
                        }`}>
                          {ref.state}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Status */}
              <div className="mt-6 p-3 bg-gray-50 rounded-lg">
                <p className="text-sm text-gray-700">
                  <span className="font-medium">Status:</span>{' '}
                  {isReady ? 'Connected' : 'Loading...'}
                </p>
              </div>
            </div>

            {/* Usage Instructions */}
            <div className="bg-white rounded-lg shadow p-6 mt-6">
              <h3 className="text-lg font-semibold mb-3">How to Use</h3>
              <ol className="list-decimal list-inside text-sm text-gray-700 space-y-2">
                <li>Click on any cell in spreadsheet</li>
                <li>Click &quot;Create Reference&quot; to create a data link</li>
                <li>References will be listed in panel</li>
                <li>Change cell values and click &quot;Refresh&quot; to update</li>
                <li>Click delete icon to remove a reference</li>
              </ol>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
