/**
 * File Operations Hook
 *
 * Provides save/load functionality for Word documents and Excel workbooks.
 */

'use client';

import { useCallback, useState } from 'react';
import { getFileAdapter } from '@/adapters';
import type { FileAdapter } from '@/adapters/filesystem';

/**
 * Supported file types for save/load operations
 */
export type FileType = 'word' | 'excel';

/**
 * File save/load state
 */
interface FileOperationsState {
  isSaving: boolean;
  isLoading: boolean;
  error: string | null;
}

/**
 * File save result
 */
export interface SaveResult {
  success: boolean;
  filename?: string;
  error?: string;
}

/**
 * File load result
 */
export interface LoadResult<T = any> {
  success: boolean;
  data?: T;
  filename?: string;
  error?: string;
}

/**
 * Use file operations hook
 */
export function useFileOperations() {
  const [state, setState] = useState<FileOperationsState>({
    isSaving: false,
    isLoading: false,
    error: null,
  });
  const [fileAdapter, setFileAdapter] = useState<FileAdapter | null>(null);

  // Initialize file adapter
  const initAdapter = useCallback(async () => {
    if (!fileAdapter) {
      const adapter = await getFileAdapter();
      setFileAdapter(adapter);
      return adapter;
    }
    return fileAdapter;
  }, [fileAdapter]);

  /**
   * Save Word document as JSON file
   */
  const saveWordDocument = useCallback(async (
    content: any,
    defaultFilename: string = 'document.json'
  ): Promise<SaveResult> => {
    try {
      setState(prev => ({ ...prev, isSaving: true, error: null }));

      const adapter = await initAdapter();

      // Prepare JSON content
      const jsonContent = JSON.stringify(content, null, 2);
      const blob = new Blob([jsonContent], { type: 'application/json' });

      // Trigger download
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = defaultFilename.endsWith('.json') ? defaultFilename : `${defaultFilename}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setState(prev => ({ ...prev, isSaving: false }));
      return { success: true, filename: defaultFilename };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to save document';
      setState(prev => ({ ...prev, isSaving: false, error: errorMessage }));
      return { success: false, error: errorMessage };
    }
  }, [initAdapter]);

  /**
   * Save Excel workbook as XLSX file
   */
  const saveExcelWorkbook = useCallback(async (
    workbookData: any,
    defaultFilename: string = 'workbook.json'
  ): Promise<SaveResult> => {
    try {
      setState(prev => ({ ...prev, isSaving: true, error: null }));

      const adapter = await initAdapter();

      // For now, save as JSON (Univer format)
      // TODO: Implement actual XLSX export using Univer's export functionality
      const jsonContent = JSON.stringify(workbookData, null, 2);
      const blob = new Blob([jsonContent], { type: 'application/json' });

      // Trigger download
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = defaultFilename.endsWith('.json') ? defaultFilename : `${defaultFilename}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setState(prev => ({ ...prev, isSaving: false }));
      return { success: true, filename: defaultFilename };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to save workbook';
      setState(prev => ({ ...prev, isSaving: false, error: errorMessage }));
      return { success: false, error: errorMessage };
    }
  }, [initAdapter]);

  /**
   * Load Word document from file
   */
  const loadWordDocument = useCallback(async (): Promise<LoadResult> => {
    try {
      setState(prev => ({ ...prev, isLoading: true, error: null }));

      const adapter = await initAdapter();

      // Open file dialog
      const files = await adapter.openFileDialog({
        filters: [
          { name: 'JSON Files', extensions: ['json'] },
          { name: 'All Files', extensions: ['*'] },
        ],
        multiple: false,
      });

      if (!files || files.length === 0) {
        setState(prev => ({ ...prev, isLoading: false }));
        return { success: false };
      }

      // Read file content
      const fileUrl = files[0];
      const response = await fetch(fileUrl);
      const content = await response.json();

      // Validate content structure (basic check)
      if (!content || typeof content !== 'object') {
        throw new Error('Invalid file format');
      }

      setState(prev => ({ ...prev, isLoading: false }));
      return { success: true, data: content };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to load document';
      setState(prev => ({ ...prev, isLoading: false, error: errorMessage }));
      return { success: false, error: errorMessage };
    }
  }, [initAdapter]);

  /**
   * Load Excel workbook from file
   */
  const loadExcelWorkbook = useCallback(async (): Promise<LoadResult> => {
    try {
      setState(prev => ({ ...prev, isLoading: true, error: null }));

      const adapter = await initAdapter();

      // Open file dialog
      const files = await adapter.openFileDialog({
        filters: [
          { name: 'JSON Files', extensions: ['json'] },
          { name: 'All Files', extensions: ['*'] },
        ],
        multiple: false,
      });

      if (!files || files.length === 0) {
        setState(prev => ({ ...prev, isLoading: false }));
        return { success: false };
      }

      // Read file content
      const fileUrl = files[0];
      const response = await fetch(fileUrl);
      const content = await response.json();

      // Validate content structure (basic check)
      if (!content || typeof content !== 'object') {
        throw new Error('Invalid file format');
      }

      setState(prev => ({ ...prev, isLoading: false }));
      return { success: true, data: content };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to load workbook';
      setState(prev => ({ ...prev, isLoading: false, error: errorMessage }));
      return { success: false, error: errorMessage };
    }
  }, [initAdapter]);

  /**
   * Clear error state
   */
  const clearError = useCallback(() => {
    setState(prev => ({ ...prev, error: null }));
  }, []);

  return {
    state,
    saveWordDocument,
    saveExcelWorkbook,
    loadWordDocument,
    loadExcelWorkbook,
    clearError,
  };
}
