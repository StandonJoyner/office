// Tauri File System Adapter Implementation
import type {
  FileAdapter,
  FileInfo,
  DocxContent,
  XlsxContent,
  FileWatcherCallback,
  UnwatchFn,
  OpenFileDialogOptions,
  SaveFileDialogOptions,
  FileFilter,
  FileWatcherEvent,
} from '../filesystem';
import { isOfficeFile, FILE_FILTERS } from '../filesystem';

/**
 * Tauri File System Adapter
 * 使用 Tauri 的 IPC 命令与 Rust 后端通信
 */

// 辅助函数：获取路径分隔符
function getPathSeparator(): string {
  if (typeof navigator !== 'undefined') {
    return navigator.platform.startsWith('Win') ? '\\' : '/';
  }
  return '/';
}

// Tauri API 类型声明
interface TauriInvokeError {
  message: string;
}

interface TauriFileMetadata {
  size: number;
  modified: number;
  accessed: number;
  created: number;
  readonly: boolean;
  is_file: boolean;
  is_dir: boolean;
  is_symlink: boolean;
}

interface TauriFileEntry {
  path: string;
  children: TauriFileEntry[];
  metadata: TauriFileMetadata;
}

// Tauri invoke helper with error handling
async function invoke<T>(command: string, args?: any): Promise<T> {
  try {
    // @ts-ignore - Tauri global
    const { invoke } = await import('@tauri-apps/api/core');
    return await invoke(command, args);
  } catch (e) {
    const error = e as TauriInvokeError;
    throw new Error(`Tauri invoke failed (${command}): ${error.message}`);
  }
}

class TauriFileAdapter implements FileAdapter {
  private watchers: Map<number, { path: string; callback: FileWatcherCallback }> = new Map();
  private watcherIdCounter: number = 0;
  private isWatching: Set<string> = new Set();

  // 基础文件操作 - Tauri 实现
  async readFile(path: string): Promise<Uint8Array> {
    const content = await invoke<number[]>('read_file', { path });
    return new Uint8Array(content);
  }

  async readFileAsText(path: string): Promise<string> {
    return await invoke<string>('read_text_file', { path });
  }

  async writeFile(path: string, data: Uint8Array | string): Promise<void> {
    if (typeof data === 'string') {
      await invoke('write_text_file', { path, contents: data });
    } else {
      await invoke('write_file', { path, contents: Array.from(data) });
    }
  }

  async deleteFile(path: string): Promise<void> {
    await invoke('remove_file', { path });
  }

  async exists(path: string): Promise<boolean> {
    return await invoke<boolean>('exists', { path });
  }

  // 目录操作 - Tauri 实现
  async listDirectory(path: string): Promise<FileInfo[]> {
    const entries = await invoke<TauriFileEntry[]>('read_dir', { path, recursive: false });
    return entries.map(entry => this.convertToFileInfo(entry));
  }

  async createDirectory(path: string): Promise<void> {
    await invoke('create_dir', { path, recursive: true });
  }

  async deleteDirectory(path: string): Promise<void> {
    await invoke('remove_dir', { path, recursive: true });
  }

  // 路径操作 - 使用 Tauri 的 path 模块
  async joinPath(...paths: string[]): Promise<string> {
    // @ts-ignore - Tauri API
    const { join } = await import('@tauri-apps/api/path');
    return await join(...paths);
  }

  async dirname(path: string): Promise<string> {
    // @ts-ignore - Tauri API
    const { dirname } = await import('@tauri-apps/api/path');
    return await dirname(path);
  }

  async basename(path: string): Promise<string> {
    // @ts-ignore - Tauri API
    const { basename } = await import('@tauri-apps/api/path');
    return await basename(path);
  }

  async normalize(path: string): Promise<string> {
    // @ts-ignore - Tauri API
    const { normalize } = await import('@tauri-apps/api/path');
    return await normalize(path);
  }

  // 同步版本（使用简单的路径处理作为回退）
  joinPathSync(...paths: string[]): string {
    return paths.join(getPathSeparator());
  }

  dirnameSync(path: string): string {
    const parts = path.split(/[/\\]/);
    parts.pop();
    return parts.join(getPathSeparator()) || getPathSeparator();
  }

  normalizeSync(path: string): string {
    return path.replace(/[/\\]+/g, getPathSeparator());
  }

  private _getBasenameSync(path: string): string {
    const parts = path.split(/[/\\]/);
    return parts[parts.length - 1] || '';
  }

  // Office 文件特定操作 - Tauri 实现
  async parseDocx(path: string): Promise<DocxContent> {
    return await invoke<DocxContent>('parse_docx', { path });
  }

  async parseXlsx(path: string): Promise<XlsxContent> {
    return await invoke<XlsxContent>('parse_xlsx', { path });
  }

  async saveDocx(path: string, content: DocxContent): Promise<void> {
    await invoke('save_docx', { path, content });
  }

  async saveXlsx(path: string, content: XlsxContent): Promise<void> {
    await invoke('save_xlsx', { path, content });
  }

  // 文件监听 - Tauri 实现
  watchFile(path: string, callback: FileWatcherCallback): UnwatchFn {
    const id = this.watcherIdCounter++;
    this.watchers.set(id, { path, callback });

    this.startWatcher(path).catch(err => {
      console.error('Failed to start file watcher:', err);
    });

    return () => {
      this.watchers.delete(id);
      // 如果没有其他监听器，停止监听
      const hasOtherWatchers = Array.from(this.watchers.values()).some(w => w.path === path);
      if (!hasOtherWatchers) {
        this.stopWatcher(path);
      }
    };
  }

  watchDirectory(path: string, callback: FileWatcherCallback): UnwatchFn {
    const id = this.watcherIdCounter++;
    this.watchers.set(id, { path, callback });

    this.startDirectoryWatcher(path).catch(err => {
      console.error('Failed to start directory watcher:', err);
    });

    return () => {
      this.watchers.delete(id);
      const hasOtherWatchers = Array.from(this.watchers.values()).some(w => w.path === path);
      if (!hasOtherWatchers) {
        this.stopWatcher(path);
      }
    };
  }

  private async startWatcher(path: string): Promise<void> {
    if (this.isWatching.has(path)) return;
    this.isWatching.add(path);

    try {
      await invoke('watch_file', {
        path,
        // 监听器 ID 用于停止
        watchId: `watch-${path}`,
      });
    } catch (e) {
      this.isWatching.delete(path);
      throw e;
    }
  }

  private async startDirectoryWatcher(path: string): Promise<void> {
    if (this.isWatching.has(path)) return;
    this.isWatching.add(path);

    try {
      await invoke('watch_directory', {
        path,
        watchId: `watch-dir-${path}`,
      });
    } catch (e) {
      this.isWatching.delete(path);
      throw e;
    }
  }

  private async stopWatcher(path: string): Promise<void> {
    this.isWatching.delete(path);
    try {
      await invoke('unwatch', { watchId: `watch-${path}` });
      await invoke('unwatch', { watchId: `watch-dir-${path}` });
    } catch {
      // 忽略错误
    }
  }

  // 桌面特有功能 - Tauri 实现
  async selectFile(filters?: FileFilter[]): Promise<string | null> {
    const files = await this.openFileDialog({
      filters: filters || [FILE_FILTERS.all],
      multiple: false,
    });
    return files[0] || null;
  }

  async selectDirectory(): Promise<string | null> {
    try {
      // @ts-ignore - Tauri API
      const { open } = await import('@tauri-apps/plugin-dialog');
      const selected = await open({
        directory: true,
      });
      return selected as string | null;
    } catch (e) {
      throw new Error(`Failed to select directory: ${(e as Error).message}`);
    }
  }

  async showInFolder(path: string): Promise<void> {
    try {
      // @ts-ignore - Tauri API
      const { showItemInFolder } = await import('@tauri-apps/api/opener');
      await showItemInFolder(path);
    } catch (e) {
      throw new Error(`Failed to show in folder: ${(e as Error).message}`);
    }
  }

  async openFileDialog(options: OpenFileDialogOptions): Promise<string[]> {
    try {
      // @ts-ignore - Tauri API
      const { open } = await import('@tauri-apps/plugin-dialog');
      const selected = await open({
        multiple: options.multiple || false,
        directory: options.directory || false,
        filters: options.filters?.map(f => ({
          name: f.name,
          extensions: f.extensions,
        })),
        defaultPath: options.defaultPath,
      });
      if (Array.isArray(selected)) {
        return selected;
      }
      return selected ? [selected] : [];
    } catch (e) {
      throw new Error(`Failed to open file dialog: ${(e as Error).message}`);
    }
  }

  async saveFileDialog(options: SaveFileDialogOptions): Promise<string | null> {
    try {
      // @ts-ignore - Tauri API
      const { save } = await import('@tauri-apps/plugin-dialog');
      const selected = await save({
        filters: options.filters?.map(f => ({
          name: f.name,
          extensions: f.extensions,
        })),
        defaultPath: options.defaultPath,
      });
      return selected as string | null;
    } catch (e) {
      throw new Error(`Failed to open save dialog: ${(e as Error).message}`);
    }
  }

  // 获取系统路径 - Tauri 实现
  async getAppDataPath(): Promise<string> {
    try {
      // @ts-ignore - Tauri API
      const { appDataDir } = await import('@tauri-apps/api/path');
      return await appDataDir();
    } catch (e) {
      return '/app-data';
    }
  }

  async getDocumentsPath(): Promise<string> {
    try {
      // @ts-ignore - Tauri API
      const { documentDir } = await import('@tauri-apps/api/path');
      return await documentDir();
    } catch (e) {
      return '/documents';
    }
  }

  async getDownloadsPath(): Promise<string> {
    try {
      // @ts-ignore - Tauri API
      const { downloadDir } = await import('@tauri-apps/api/path');
      return await downloadDir();
    } catch (e) {
      return '/downloads';
    }
  }

  async getHomePath(): Promise<string> {
    try {
      // @ts-ignore - Tauri API
      const { homeDir } = await import('@tauri-apps/api/path');
      return await homeDir();
    } catch (e) {
      return '/home';
    }
  }

  // Helper methods
  private convertToFileInfo(entry: TauriFileEntry): FileInfo {
    return {
      name: this._getBasenameSync(entry.path),
      path: entry.path,
      isDirectory: entry.metadata.is_dir,
      size: entry.metadata.size,
      modifiedAt: entry.metadata.modified,
      createdAt: entry.metadata.created,
    };
  }

  // 处理来自 Tauri 的文件监听事件
  handleFileEvent(event: FileWatcherEvent): void {
    for (const watcher of this.watchers.values()) {
      if (watcher.path === event.path || (watcher.path.includes(event.path))) {
        watcher.callback(event);
      }
    }
  }
}

// 导出单例实例
export let tauriFileAdapter: TauriFileAdapter | null = null;

/**
 * 创建 Tauri 文件系统适配器实例
 * 只在 Tauri 环境中可用
 */
export async function createTauriFileAdapter(): Promise<TauriFileAdapter> {
  if (!tauriFileAdapter) {
    tauriFileAdapter = new TauriFileAdapter();
  }
  return tauriFileAdapter;
}

/**
 * 检查是否在 Tauri 环境中运行
 */
export function isTauriEnvironment(): boolean {
  return typeof window !== 'undefined' && (window as any).__TAURI__ !== undefined;
}
