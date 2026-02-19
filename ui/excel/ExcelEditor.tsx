'use client';

/**
 * Excel Editor Component
 * Based on official Univer example: https://github.com/dream-num/univer/blob/master/examples/src/sheets/main.ts
 */

import React, { useRef, useState, useCallback, useEffect } from 'react';
import { LocaleType, Univer, UniverInstanceType } from '@univerjs/core';
import { FUniver } from '@univerjs/core/facade';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui';
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin } from '@univerjs/ui';
import { eventBus, Events } from '@/core/eventBus';

// Import facade extensions
import '@univerjs/sheets/facade';
import '@univerjs/ui/facade';
import '@univerjs/docs-ui/facade';
import '@univerjs/sheets-ui/facade';
import '@univerjs/engine-formula/facade';
import '@univerjs/sheets-formula/facade';
import '@univerjs/sheets-numfmt/facade';
import '@univerjs/sheets-formula-ui/facade';

// Import Univer styles
import '@univerjs/design/lib/index.css';
import '@univerjs/ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
import '@univerjs/sheets-formula-ui/lib/index.css';

export interface ExcelEditorProps {
  initialData?: any;
  readOnly?: boolean;
  enableDataSource?: boolean;
  onWorkbookLoaded?: (workbook: any) => void;
  onCellSelected?: (sheetId: string, row: number, column: number) => void;
  onRangeSelected?: (sheetId: string, range: { startRow: number; startCol: number; endRow: number; endCol: number }) => void;
  onDataChange?: (event: any) => void;
  theme?: { primaryColor?: string; fontSize?: number };
  height?: string | number;
  className?: string;
  fileId?: string;
  fileName?: string;
}

export const ExcelEditor = React.forwardRef<ExcelEditorRefType, ExcelEditorProps>(({
  initialData,
  enableDataSource = true,
  onWorkbookLoaded,
  onCellSelected,
  onRangeSelected,
  onDataChange,
  height = '600px',
  className = '',
  fileName = 'Untitled.xlsx',
}, forwardedRef) => {
  ExcelEditor.displayName = 'ExcelEditor';

  const containerRef = useRef<HTMLDivElement>(null);
  const univerAPIRef = useRef<FUniver | null>(null);
  const workbookRef = useRef<any>(null);
  const univerInstanceRef = useRef<Univer | null>(null);
  const initializedRef = useRef(false);
  const [isReady, setIsReady] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Use a ref to track if we're on the client and generate ID only once
  const containerIdRef = useRef<string>('');
  const [mounted, setMounted] = useState(false);

  // Generate container ID only after mount (client-side only)
  useEffect(() => {
    containerIdRef.current = `univer-${Math.random().toString(36).slice(2, 10)}`;
    setMounted(true);
  }, []);

  // Setup listeners callback
  const setupDataSourceListeners = useCallback((api: FUniver) => {
    api.addEvent('sheet.operation' as any, (params: any) => {
      if (params.type !== 'SET_RANGE_VALUES' || !params.range) return;
      const activeWorkbook = api.getActiveWorkbook();
      const activeSheet = activeWorkbook?.getActiveSheet();
      if (!activeSheet) return;

      const rangeValue = (params.range as any).getValue?.();
      if (rangeValue === undefined) return;

      eventBus.emit(Events.DataChanged, {
        referenceId: crypto.randomUUID(),
        sheetId: activeSheet.getSheetId(),
        oldValue: null,
        newValue: rangeValue,
        timestamp: Date.now(),
        source: 'manual',
      });

      onDataChange?.({
        sheetId: activeSheet.getSheetId(),
        range: params.range,
        oldValue: null,
        newValue: rangeValue,
      });
    });

    api.addEvent('sheet.operation' as any, (event: any) => {
      if (event.type !== 'RANGE_SELECTION') return;
      const { selections } = event;
      if (!selections?.length) return;

      const selection = selections[0];
      const { sheetId, range } = selection;
      if (range) {
        const { startRow, endRow, startColumn, endColumn } = range;
        if (endRow !== undefined && endColumn !== undefined) {
          onRangeSelected?.(sheetId, { startRow, startCol: startColumn, endRow, endCol: endColumn });
        } else {
          onCellSelected?.(sheetId, startRow, startColumn);
        }
      }
    });
  }, [onCellSelected, onRangeSelected, onDataChange]);

  // Initialize Univer once on mount
  useEffect(() => {
    // Only run on client after mounted and not already initialized
    if (!mounted || initializedRef.current) {
      return;
    }

    const containerId = containerIdRef.current;
    if (!containerId) {
      return;
    }

    const init = async () => {
      // Wait for the container element to be in the DOM
      const containerEl = document.getElementById(containerId);
      if (!containerEl) {
        console.log('[ExcelEditor] Container element not found in DOM, retrying...');
        // Retry after a short delay
        setTimeout(init, 50);
        return;
      }

      initializedRef.current = true;
      console.log('[ExcelEditor] Initializing...');
      console.log('[ExcelEditor] Container dimensions:', containerEl.offsetWidth, 'x', containerEl.offsetHeight);

      try {
        // Import locale data
        const localeData = await import('@univerjs/preset-sheets-core/locales/en-US');

        // Create Univer instance following official example pattern
        const univer = new Univer({
          locale: LocaleType.EN_US,
          locales: {
            [LocaleType.EN_US]: localeData,
          },
        });

        // Register plugins using array syntax (official pattern)
        univer.registerPlugins([
          [UniverDocsPlugin],
          [UniverRenderEnginePlugin],
          [UniverUIPlugin, {
            container: containerId,
          }],
          [UniverDocsUIPlugin],
          [UniverSheetsPlugin],
          [UniverSheetsUIPlugin],
          [UniverFormulaEnginePlugin],
          [UniverSheetsFormulaPlugin],
          [UniverSheetsFormulaUIPlugin],
          [UniverSheetsNumfmtPlugin],
          [UniverSheetsNumfmtUIPlugin],
        ]);

        univerInstanceRef.current = univer;

        // Create API
        const univerAPI = FUniver.newAPI(univer);
        univerAPIRef.current = univerAPI;

        console.log('[ExcelEditor] Univer created, container children:', containerEl.children.length);

        // Create a default sheet
        const sheetData = {
          id: 'workbook-' + Date.now(),
          sheetOrder: ['sheet-01'],
          name: 'Workbook',
          appVersion: '1.0.0',
          sheets: {
            'sheet-01': {
              id: 'sheet-01',
              name: 'Sheet1',
              rowCount: 100,
              columnCount: 26,
              cellData: {},
            },
          },
        };

        const unit = univer.createUnit(UniverInstanceType.UNIVER_SHEET, initialData || sheetData);
        console.log('[ExcelEditor] Sheet unit created:', unit?.getUnitId?.());

        // Check DOM after sheet creation
        setTimeout(() => {
          console.log('[ExcelEditor] Container after sheet (delayed):', containerEl.children.length);
          console.log('[ExcelEditor] Container children:', Array.from(containerEl.children).map(c => c.tagName + '.' + c.className.split(' ')[0]));
          const canvas = containerEl.querySelector('canvas');
          console.log('[ExcelEditor] Canvas found:', !!canvas, canvas ? `${canvas.width}x${canvas.height}` : 'none');
        }, 100);

        if (enableDataSource) {
          setupDataSourceListeners(univerAPI);
        }

        const activeWorkbook = univerAPI.getActiveWorkbook();
        workbookRef.current = activeWorkbook;

        setIsReady(true);
        onWorkbookLoaded?.(activeWorkbook);
        console.log('[ExcelEditor] Ready!');
      } catch (err) {
        console.error('[ExcelEditor] Error:', err);
        setError(err instanceof Error ? err.message : 'Failed to initialize');
      }
    };

    init();
  }, [enableDataSource, initialData, onWorkbookLoaded, setupDataSourceListeners, mounted]);

  // Expose methods
  React.useImperativeHandle(forwardedRef, () => ({
    getWorkbook: () => workbookRef.current,
    getActiveSheet: () => workbookRef.current?.getActiveSheet(),
    getCellValue: async (sheetId: string, row: number, column: number) => {
      const api = univerAPIRef.current;
      const sheet = api?.getActiveWorkbook()?.getSheetBySheetId?.(sheetId);
      return sheet ? (sheet as any).getRangeByRow?.(row, column, 1, 1)?.getValue() ?? null : null;
    },
    setCellValue: async (sheetId: string, row: number, column: number, value: any) => {
      const api = univerAPIRef.current;
      const sheet = api?.getActiveWorkbook()?.getSheetBySheetId?.(sheetId);
      (sheet as any)?.getRangeByRow?.(row, column, 1, 1)?.setValue(value);
    },
    getWorkbookData: () => univerAPIRef.current?.getActiveWorkbook()?.save() ?? null,
    getCellValueByA1: async (a1: string) => {
      const sheet = univerAPIRef.current?.getActiveWorkbook()?.getActiveSheet();
      return sheet?.getRange(a1)?.getValue() ?? null;
    },
    setCellValueByA1: async (a1: string, value: any) => {
      const sheet = univerAPIRef.current?.getActiveWorkbook()?.getActiveSheet();
      sheet?.getRange(a1)?.setValue(value);
    },
    getSelection: () => {
      const sheet = univerAPIRef.current?.getActiveWorkbook()?.getActiveSheet();
      return (sheet as any)?.getSelections?.() ?? null;
    },
    getAPI: () => univerAPIRef.current,
  }));

  if (error) {
    return (
      <div className={`excel-editor-error ${className}`} style={{ height }}>
        <div className="p-4">
          <h3 className="text-lg font-semibold text-red-800">Failed to load Excel editor</h3>
          <p className="text-sm text-red-600 mt-2">{error}</p>
        </div>
      </div>
    );
  }

  return (
    <div
      className={`border border-slate-200 rounded-lg overflow-hidden bg-white flex flex-col ${className}`}
      style={{ height: typeof height === 'number' ? `${height}px` : height }}
    >
      <div className="flex items-center gap-2 p-2 border-b border-slate-200 bg-slate-50 flex-shrink-0">
        <span className="text-sm text-slate-600 font-medium">{fileName}</span>
        <div className="w-px h-6 bg-slate-300" />
        <span className={`text-sm ${isReady ? 'text-green-600' : 'text-slate-500'}`}>
          {isReady ? 'Ready' : 'Loading...'}
        </span>
      </div>
      <div
        ref={containerRef}
        id={mounted ? containerIdRef.current : undefined}
        className="flex-1"
        style={{ height: '100%', width: '100%', position: 'relative' }}
      />
    </div>
  );
});

export interface ExcelEditorRefType {
  getWorkbook: () => any;
  getActiveSheet: () => any;
  getCellValue: (sheetId: string, row: number, column: number) => Promise<any>;
  setCellValue: (sheetId: string, row: number, column: number, value: any) => Promise<void>;
  getWorkbookData: () => any;
  getCellValueByA1: (a1: string) => Promise<any>;
  setCellValueByA1: (a1: string, value: any) => Promise<void>;
  getSelection: () => any;
  getAPI: () => FUniver | null;
}

export function columnToLetter(index: number): string {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

export function cellToA1(row: number, col: number): string {
  return `${columnToLetter(col)}${row + 1}`;
}

export function a1ToCell(a1: string): { row: number; col: number } {
  const match = a1.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { row: 0, col: 0 };
  let col = 0;
  for (let i = 0; i < match[1].length; i++) {
    col = col * 26 + (match[1].charCodeAt(i) - 64);
  }
  return { row: parseInt(match[2], 10) - 1, col: col - 1 };
}
