// File System Adapter Interface for Web/Desktop abstraction
import type { WordDocument, ExcelDocument } from '../core/types';

export interface FileAdapter {
  // 基础文件操作
  readFile(path: string): Promise<Uint8Array>;
  readFileAsText(path: string): Promise<string>;
  writeFile(path: string, data: Uint8Array | string): Promise<void>;
  deleteFile(path: string): Promise<void>;
  exists(path: string): Promise<boolean>;

  // 目录操作
  listDirectory(path: string): Promise<FileInfo[]>;
  createDirectory(path: string): Promise<void>;
  deleteDirectory(path: string): Promise<void>;

  // 路径操作 (支持异步以适应 Tauri API)
  joinPath(...paths: string[]): Promise<string> | string;
  dirname(path: string): Promise<string> | string;
  basename(path: string): Promise<string> | string;
  normalize(path: string): Promise<string> | string;

  // Office 文件特定操作
  parseDocx(path: string): Promise<DocxContent>;
  parseXlsx(path: string): Promise<XlsxContent>;
  saveDocx(path: string, content: DocxContent): Promise<void>;
  saveXlsx(path: string, content: XlsxContent): Promise<void>;

  // 文件监听
  watchFile(path: string, callback: FileWatcherCallback): UnwatchFn;
  watchDirectory(path: string, callback: FileWatcherCallback): UnwatchFn;

  // 桌面特有功能
  selectFile(filters?: FileFilter[]): Promise<string | null>;
  selectDirectory(): Promise<string | null>;
  showInFolder(path: string): Promise<void>;
  openFileDialog(options: OpenFileDialogOptions): Promise<string[]>;
  saveFileDialog(options: SaveFileDialogOptions): Promise<string | null>;

  // 获取系统路径
  getAppDataPath(): Promise<string>;
  getDocumentsPath(): Promise<string>;
  getDownloadsPath(): Promise<string>;
  getHomePath(): Promise<string>;
}

// Type Definitions
export interface FileInfo {
  name: string;
  path: string;
  isDirectory: boolean;
  size: number;
  modifiedAt: number;
  createdAt: number;
}

export interface DocxContent {
  text: string;
  paragraphs: DocxParagraph[];
  tables: DocxTable[];
  metadata: {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    createdAt?: Date;
    modifiedAt?: Date;
  };
}

export interface DocxParagraph {
  id: string;
  text: string;
  style?: string;
  runs: DocxRun[];
}

export interface DocxRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  size?: number;
}

export interface DocxTable {
  id: string;
  rows: DocxRow[];
  style?: string;
}

export interface DocxRow {
  cells: DocxCell[];
}

export interface DocxCell {
  text: string;
  spans?: number;
}

export interface XlsxContent {
  sheets: XlsxSheet[];
  metadata: {
    title?: string;
    author?: string;
    createdAt?: Date;
    modifiedAt?: Date;
  };
}

export interface XlsxSheet {
  id: string;
  name: string;
  cells: Map<string, XlsxCell>;
  namedRanges?: XlsxNamedRange[];
  charts?: XlsxChart[];
}

export interface XlsxCell {
  value: any;
  formula?: string;
  format?: string;
  style?: XlsxCellStyle;
}

export interface XlsxCellStyle {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontColor?: string;
  backgroundColor?: string;
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
  };
}

export interface XlsxNamedRange {
  name: string;
  range: string;
  sheetId: string;
}

export interface XlsxChart {
  id: string;
  type: 'line' | 'bar' | 'pie' | 'scatter';
  title?: string;
  dataRange: string;
  sheetId: string;
}

export interface FileFilter {
  name: string;
  extensions: string[];
}

export interface OpenFileDialogOptions {
  title?: string;
  defaultPath?: string;
  buttonLabel?: string;
  filters?: FileFilter[];
  multiple?: boolean;
  directory?: boolean;
}

export interface SaveFileDialogOptions {
  title?: string;
  defaultPath?: string;
  buttonLabel?: string;
  filters?: FileFilter[];
}

export type FileWatcherCallback = (event: FileWatcherEvent) => void;

export interface FileWatcherEvent {
  type: 'created' | 'modified' | 'deleted' | 'renamed';
  path: string;
  oldPath?: string; // For renamed events
  timestamp: number;
}

export type UnwatchFn = () => void;

// Platform Detection
export function isDesktop(): boolean {
  if (typeof window === 'undefined') return false;
  return (window as any).__TAURI__ !== undefined;
}

// For use in other modules that also need isDesktop
export { isDesktop as _isDesktop };

export function getPlatform(): 'windows' | 'macos' | 'linux' | 'web' {
  if (typeof window === 'undefined') return 'web';
  if (isDesktop()) {
    // @ts-ignore - Tauri API
    return window.__TAURI__.os.platform();
  }
  if (navigator.platform.startsWith('Win')) return 'windows';
  if (navigator.platform.startsWith('Mac')) return 'macos';
  if (navigator.platform.startsWith('Linux')) return 'linux';
  return 'web';
}

// Path separator helper
export function pathSeparator(): string {
  return getPlatform() === 'windows' ? '\\' : '/';
}

// File extension helpers
export function getExtension(path: string): string {
  const parts = path.split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : '';
}

export function isDocxFile(path: string): boolean {
  return getExtension(path) === 'docx';
}

export function isXlsxFile(path: string): boolean {
  return getExtension(path) === 'xlsx';
}

export function isOfficeFile(path: string): boolean {
  const ext = getExtension(path);
  return ['docx', 'xlsx', 'pptx', 'doc', 'xls', 'ppt'].includes(ext);
}

// Default file filters
export const FILE_FILTERS = {
  all: { name: 'All Files', extensions: ['*'] },
  office: {
    name: 'Office Documents',
    extensions: ['docx', 'xlsx', 'pptx', 'doc', 'xls'],
  },
  word: {
    name: 'Word Documents',
    extensions: ['docx', 'doc'],
  },
  excel: {
    name: 'Excel Spreadsheets',
    extensions: ['xlsx', 'xls'],
  },
  images: {
    name: 'Images',
    extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'],
  },
};

// MIME type helper
export function getMimeType(extension: string): string {
  const types: Record<string, string> = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    doc: 'application/msword',
    xls: 'application/vnd.ms-excel',
    ppt: 'application/vnd.ms-powerpoint',
    pdf: 'application/pdf',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
    svg: 'image/svg+xml',
    txt: 'text/plain',
    json: 'application/json',
    xml: 'application/xml',
  };
  return types[extension.toLowerCase()] || 'application/octet-stream';
}
