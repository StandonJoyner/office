'use client';

import { useState, useRef, useMemo, useEffect, useCallback } from 'react';
import { WordEditor, type WordEditorRefType } from '@/ui/word';
import { TraceabilityPanel } from '@/ui/common';
import { DataUpdateNotification } from '@/ui/components/DataUpdateNotification';
import { ExcelEditor, type ExcelEditorRefType } from '@/ui/excel';
import { createDataLinkManager, createDataSyncEngine } from '@/core';
import { createTraceabilityService, type TraceabilityService } from '@/core/trace/TraceabilityService';
import { eventBus, Events } from '@/core/eventBus';
import { DataReference, type DataSource } from '@/core/types';
import { createUniverWorkbookAdapter, type IFUniverWorkbook, type IFUniverAPI } from '@/core/editors/UniverWorkbookAdapter';

/**
 * Cell Selection State
 */
interface CellSelection {
  sheetId: string;
  sheetName: string;
  row: number;
  column: number;
  value: any;
  address: string;
}

export default function IntegratedPage() {
  const excelEditorRef = useRef<ExcelEditorRefType>(null);
  const wordEditorRef = useRef<WordEditorRefType>(null);
  const dataLinkManagerRef = useRef(createDataLinkManager());
  const syncEngineRef = useRef(
    createDataSyncEngine(
      dataLinkManagerRef.current,
      (ref) => dataLinkManagerRef.current!.resolveReference(ref)
    )
  );
  const traceabilityServiceRef = useRef<TraceabilityService | null>(null);
  const workbookAdapterRef = useRef<any>(null);

  // State
  const [selectedReference, setSelectedReference] = useState<DataReference | null>(null);
  const [selectedCell, setSelectedCell] = useState<CellSelection | null>(null);
  const [isInserting, setIsInserting] = useState(false);
  const [references, setReferences] = useState<DataReference[]>([]);
  const [documentContent, setDocumentContent] = useState<any>({
    type: 'doc',
    content: [
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'Office Suite 集成演示' },
        ],
      },
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: '这是 Word 和 Excel 之间的数据引用集成演示。' },
        ],
      },
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: '从右侧 Excel 编辑器中选择单元格，然后点击 "插入引用" 按钮在 Word 文档中创建数据引用。' },
        ],
      },
    ],
  });

  const documentId = 'integrated-doc-1';

  // Create traceability service with proper loaders
  traceabilityServiceRef.current = useMemo(() => {
    return createTraceabilityService(
      (docId) => {
        // Get document content
        return documentId === docId ? documentContent : null;
      },
      (docId) => {
        // Get document references
        return dataLinkManagerRef.current.findReferencesByDocument(docId)
          .map(id => dataLinkManagerRef.current.getReference(id))
          .filter((r): r is DataReference => r !== null);
      }
    );
  }, [documentId, documentContent]);

  const traceabilityService = traceabilityServiceRef.current;

  // Load initial document references
  useEffect(() => {
    const refIds = dataLinkManagerRef.current.findReferencesByDocument(documentId);
    const refs = refIds
      .map(id => dataLinkManagerRef.current.getReference(id))
      .filter((r): r is DataReference => r !== null);
    setReferences(refs);
  }, [documentId]);

  // Listen to reference events to update state
  useEffect(() => {
    const handleReferenceCreated = (data: any) => {
      const ref = data.reference;
      if (ref && ref.target?.documentId === documentId) {
        setReferences(prev => [...prev, ref]);
        setSelectedReference(ref);
      }
    };

    const handleReferenceUpdated = (data: any) => {
      const { referenceId } = data;
      const updatedRef = dataLinkManagerRef.current.getReference(referenceId);
      if (updatedRef) {
        setReferences(prev => prev.map(r => r.id === referenceId ? updatedRef : r));
        if (selectedReference?.id === referenceId) {
          setSelectedReference(updatedRef);
        }
      }
    };


    eventBus.on(Events.ReferenceCreated, handleReferenceCreated);
    eventBus.on(Events.ReferenceUpdated, handleReferenceUpdated);

    return () => {
      eventBus.off(Events.ReferenceCreated, handleReferenceCreated);
      eventBus.off(Events.ReferenceUpdated, handleReferenceUpdated);
    };
  }, [documentId, selectedReference]);

  // Listen for Excel cell changes using workbook adapter for auto-sync
  useEffect(() => {
    if (!workbookAdapterRef.current) return;

    const handleCellChange = async (data: any) => {
      console.log('Cell change detected:', data);
      const { sheetId, row, column, value, oldValue } = data;

      // Skip if value hasn't actually changed
      if (oldValue === value) return;

      // Find all references that are affected by this cell change
      const allRefs = dataLinkManagerRef.current.getAllReferences();
      const affectedRefs: DataReference[] = [];

      for (const ref of allRefs) {
        // Check if reference is in this document
        if (ref.target.documentId !== documentId) continue;

        // Check if reference's source matches the changed cell
        if (ref.source.sheetId === sheetId) {
          const { range } = ref.source;

          // Check if the changed cell is within the reference's source range
          const refEndRow = range.endRow ?? range.startRow;
          const refEndCol = range.endCol ?? range.startCol;

          const isAffected =
            row >= range.startRow &&
            row <= refEndRow &&
            column >= range.startCol &&
            column <= refEndCol;

          if (isAffected) {
            affectedRefs.push(ref);
          }
        }
      }

      if (affectedRefs.length === 0) return;

      console.log(`Found ${affectedRefs.length} affected references for cell change`);

      // Update each affected reference
      for (const ref of affectedRefs) {
        try {
          // Use data link manager to resolve reference (gets new value from Excel)
          const newValue = await dataLinkManagerRef.current.resolveReference(ref);

          // Update reference in DataLinkManager
          dataLinkManagerRef.current.updateReference(ref.id, {
            display: {
              ...ref.display,
              value: newValue,
            },
            state: 'active',
            metadata: {
              updated: Date.now() as number,
              created: Date.now() as number,
            },
          });

          // Update local state
          setReferences(prev =>
            prev.map(r =>
              r.id === ref.id
                ? { ...r, display: { ...r.display, value: newValue }, state: 'active' as const }
                : r
            )
          );

          if (selectedReference?.id === ref.id) {
            setSelectedReference(prev =>
              prev ? { ...prev, display: { ...prev.display, value: newValue }, state: 'active' as const } : null
            );
          }

          // Emit ReferenceUpdated event to trigger editor update
          eventBus.emit(Events.ReferenceUpdated, {
            referenceId: ref.id,
            newState: { ...ref, display: { ...ref.display, value: newValue }, state: 'active' as const },
          });
        } catch (error) {
          console.error(`Failed to resolve reference ${ref.id}:`, error);

          // Mark reference as broken
          dataLinkManagerRef.current.updateReference(ref.id, {
            state: 'broken',
            metadata: {
              ...ref.metadata,
              updated: Date.now(),
            },
          });

          setReferences(prev =>
            prev.map(r =>
              r.id === ref.id ? { ...r, state: 'broken' as const } : r
            )
          );

          if (selectedReference?.id === ref.id) {
            setSelectedReference(prev =>
              prev ? { ...prev, state: 'broken' as const } : null
            );
          }
        }
      }
    };

    // Register cell change listener on workbook adapter
    workbookAdapterRef.current.on('cellChange', handleCellChange);

    return () => {
      workbookAdapterRef.current?.off('cellChange', handleCellChange);
    };
  }, [documentId, selectedReference]);

  // Handle Excel workbook loaded
  const handleWorkbookLoaded = useCallback((workbook: IFUniverWorkbook) => {
    console.log('Excel workbook loaded:', workbook);

    // Get the FUniver API from the ExcelEditor
    const api = excelEditorRef.current?.getAPI?.();
    if (workbook && api) {
      // Create adapter and set it on data link manager
      const adapter = createUniverWorkbookAdapter(workbook as any, api as any);
      workbookAdapterRef.current = adapter;
      dataLinkManagerRef.current.setExcelWorkbook(adapter);

      console.log('Univer workbook adapter created and set');
    }
  }, []);

  // Handle cell selection in Excel (single cell)
  const handleCellSelected = useCallback(async (sheetId: string, row: number, column: number) => {

    const api = excelEditorRef.current?.getAPI?.();
    if (!api) return;

    const workbook = api.getActiveWorkbook();
    if (!workbook) return;

    const sheet = workbook.getSheetBySheetId(sheetId);
    if (!sheet) return;

    const range = (sheet as any).getRangeByRow?.(row, column);
    if (!range) return;

    const value = range.getValue();
    const address = range.getAddress();

    setSelectedCell({
      sheetId,
      sheetName: (sheet as any).getName?.() || '',
      row,
      column,
      value,
      address,
    });
  }, []);

  // Handle range selection in Excel (Univer fires onRangeSelected even for single cell)
  const handleRangeSelected = useCallback(
    (sheetId: string, range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
      handleCellSelected(sheetId, range.startRow, range.startCol);
    },
    [handleCellSelected]
  );

  // Handle reference creation from Word editor
  const handleReferenceCreated = useCallback((reference: DataReference) => {
    console.log('Reference created:', reference);
    setSelectedReference(reference);
  }, []);

  // Handle document content change
  const handleContentChange = useCallback((content: any) => {
    setDocumentContent(content);
  }, []);

  // Insert reference from selected Excel cell (or current selection from API)
  const handleInsertReference = useCallback(async () => {
    if (!wordEditorRef.current) return;

    let cellInfo = selectedCell;

    // Fallback: fetch current selection from Excel API when selection event didn't fire
    if (!cellInfo) {
      const sheet = excelEditorRef.current?.getAPI?.()?.getActiveWorkbook?.()?.getActiveSheet?.();
      const selection = (sheet as any)?.getSelection?.();

      if (selection) {
        const currentCell = selection.getCurrentCell?.();
        const activeRange = selection.getActiveRange?.();
        const row = currentCell?.actualRow ?? activeRange?.getRow?.() ?? 0;
        const col = currentCell?.actualColumn ?? activeRange?.getColumn?.() ?? 0;

        // Use activeRange (FRange) for value/address - sheet.getRangeByRow may not exist on raw worksheet
        const value = activeRange?.getValue?.() ?? (sheet as any)?.getRangeByRow?.(row, col)?.getValue?.();
        const address = activeRange?.getAddress?.() ?? activeRange?.getA1Notation?.() ?? (sheet as any)?.getRangeByRow?.(row, col)?.getAddress?.() ?? `R${row + 1}C${col + 1}`;

        if (sheet) {
          cellInfo = {
            sheetId: (sheet as any).getSheetId?.() || 'sheet-01',
            sheetName: (sheet as any).getName?.() || 'Sheet1',
            row,
            column: col,
            value: value ?? null,
            address: address || `R${row + 1}C${col + 1}`,
          };
        }
      }
    }

    if (!cellInfo) {
      alert('请先在 Excel 中选择一个单元格');
      return;
    }

    setIsInserting(true);
    try {
      const dataSource: DataSource = {
        fileId: 'demo-workbook',
        fileName: 'Demo.xlsx',
        sheetId: cellInfo.sheetId,
        sheetName: cellInfo.sheetName,
        range: {
          startRow: cellInfo.row,
          startCol: cellInfo.column,
        },
        isFormula: false,
      };

      const reference: DataReference = {
        id: crypto.randomUUID(),
        type: 'cell',
        source: dataSource,
        target: {
          documentId,
          nodeId: crypto.randomUUID(),
        },
        display: {
          format: 'value',
          value: cellInfo.value,
          expression: `[Demo.xlsx!${cellInfo.sheetName}!${cellInfo.address}]`,
          tooltip: `Source: Demo.xlsx, ${cellInfo.sheetName}, ${cellInfo.address}`,
        },
        state: 'active',
        history: [],
      };

      wordEditorRef.current.insertReference(reference);
      setSelectedReference(reference);
      if (!selectedCell) setSelectedCell(cellInfo);
    } catch (error) {
      console.error('Error inserting reference:', error);
    } finally {
      setIsInserting(false);
    }
  }, [selectedCell, documentId]);

  // Refresh all references in Word (use refs so callback always has latest)
  const handleRefreshAllReferences = useCallback(async () => {
    const engine = syncEngineRef.current;
    if (engine) await engine.refreshDocument(documentId);
  }, [documentId]);

  const onRefreshAllRef = useRef<(() => Promise<void>) | null>(null);
  onRefreshAllRef.current = handleRefreshAllReferences;

  // Handle reference node insertion in editor
  const handleReferenceInserted = useCallback((reference: DataReference) => {
    console.log('Reference inserted in editor:', reference);
    // Add to references list if not already there
    setReferences(prev => {
      if (prev.find(r => r.id === reference.id)) {
        return prev;
      }
      return [...prev, reference];
    });
  }, []);

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      <DataUpdateNotification />

      <div className="container mx-auto p-6 max-w-7xl">
        {/* Header */}
        <div className="mb-6">
          <h1 className="text-3xl font-bold text-slate-900">
            Office Suite 集成演示
          </h1>
          <p className="mt-2 text-lg text-slate-600">
            Word + Excel 编辑器，支持数据引用与实时同步
          </p>
        </div>

        {/* Selected Cell Info Panel */}
        {selectedCell && (
          <div className="mb-4 p-4 bg-amber-50 rounded-lg border border-amber-200">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-sm font-semibold text-amber-900 mb-1">
                  已选中单元格
                </h3>
                <div className="text-sm text-amber-800 space-y-0.5">
                  <p><strong>工作表:</strong> {selectedCell.sheetName}</p>
                  <p><strong>位置:</strong> {selectedCell.address}</p>
                  <p><strong>值:</strong> {JSON.stringify(selectedCell.value)}</p>
                </div>
              </div>
              <button
                onClick={handleInsertReference}
                disabled={isInserting}
                className="px-4 py-2 bg-amber-600 text-white rounded-lg hover:bg-amber-700 disabled:bg-amber-400 disabled:cursor-not-allowed transition-colors"
              >
                {isInserting ? '插入中...' : '插入引用到 Word'}
              </button>
            </div>
          </div>
        )}

        {/* Main Content Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Word Editor Panel */}
          <div className="lg:col-span-1">
            <div className="bg-white rounded-lg shadow p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-semibold text-slate-900">
                  Word 文档
                </h2>
                <div className="text-sm text-slate-500">
                  {references.length} 个引用
                </div>
              </div>
              <WordEditor
                ref={wordEditorRef}
                content={documentContent}
                onChange={handleContentChange}
                onReferenceInserted={handleReferenceInserted}
                onInsertReferenceRequest={handleInsertReference}
                documentId={documentId}
                dataLinkManager={dataLinkManagerRef.current}
                syncEngine={syncEngineRef.current}
                onRefreshAll={handleRefreshAllReferences}
                onRefreshAllRef={onRefreshAllRef}
              />

              {/* Selected Reference Info */}
              {selectedReference && (
                <div className="mt-4 p-4 bg-blue-50 rounded-lg border border-blue-200">
                  <div className="flex items-center justify-between mb-2">
                    <h3 className="text-sm font-semibold text-blue-900">
                      已选引用
                    </h3>
                    <span className={`px-2 py-0.5 text-xs rounded ${
                      selectedReference.state === 'active'
                        ? 'bg-green-200 text-green-800'
                        : selectedReference.state === 'stale'
                        ? 'bg-yellow-200 text-yellow-800'
                        : selectedReference.state === 'broken'
                        ? 'bg-red-200 text-red-800'
                        : 'bg-orange-200 text-orange-800'
                    }`}>
                      {selectedReference.state}
                    </span>
                  </div>
                  <div className="space-y-1 text-sm text-blue-700">
                    <p><strong>文件:</strong> {selectedReference.source.fileName}</p>
                    <p><strong>工作表:</strong> {selectedReference.source.sheetName}</p>
                    <p><strong>单元格:</strong> ({selectedReference.source.range.startRow + 1}, {selectedReference.source.range.startCol + 1})</p>
                    <p><strong>值:</strong> {selectedReference.display.value ?? '加载中...'}</p>
                  </div>
                </div>
              )}

              {/* References List */}
              {references.length > 0 && (
                <div className="mt-4">
                  <h3 className="text-sm font-semibold text-slate-700 mb-2">
                    文档中的引用
                  </h3>
                  <div className="space-y-2 max-h-40 overflow-y-auto">
                    {references.map((ref) => (
                      <button
                        key={ref.id}
                        onClick={() => setSelectedReference(ref)}
                        className={`w-full text-left p-2 rounded border transition-colors ${
                          selectedReference?.id === ref.id
                            ? 'border-blue-500 bg-blue-50'
                            : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'
                        }`}
                      >
                        <div className="text-sm font-medium text-slate-800">
                          {ref.source.sheetName}!{ref.display.expression?.match(/\[(.*?)\]/)?.[1] || 'Unknown'}
                        </div>
                        <div className="text-xs text-slate-500">
                          {ref.display.value?.toString() ?? 'N/A'}
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Excel Editor Panel */}
          <div className="lg:col-span-1">
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-xl font-semibold text-slate-900 mb-4">
                Excel 电子表格
              </h2>
              <ExcelEditor
                ref={excelEditorRef as any}
                height="600px"
                enableDataSource={true}
                onWorkbookLoaded={handleWorkbookLoaded}
                onCellSelected={handleCellSelected}
                onRangeSelected={handleRangeSelected}
                fileId="demo-workbook"
                fileName="Demo.xlsx"
              />
              <p className="mt-4 text-sm text-slate-600">
                选择一个单元格来创建可以在 Word 文档中使用的数据引用。
              </p>
            </div>
          </div>
        </div>

        {/* Traceability Panel */}
        <div className="mt-6">
          <div className="bg-white rounded-lg shadow p-6">
            <h2 className="text-xl font-semibold text-slate-900 mb-4">
              数据溯源
            </h2>
            <TraceabilityPanel documentId={documentId} traceabilityService={traceabilityService!} />
          </div>
        </div>
      </div>
    </main>
  );
}
