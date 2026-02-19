// Web File System Adapter Implementation
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
} from '../filesystem';
import { isOfficeFile, FILE_FILTERS, getMimeType } from '../filesystem';

class WebFileAdapter implements FileAdapter {
  private watchers: Map<string, Set<FileWatcherCallback>> = new Map();
  private watcherId: number = 0;

  // 基础文件操作 - Web 实现
  async readFile(path: string): Promise<Uint8Array> {
    const response = await fetch(path);
    if (!response.ok) {
      throw new Error(`Failed to read file: ${response.statusText}`);
    }
    const buffer = await response.arrayBuffer();
    return new Uint8Array(buffer);
  }

  async readFileAsText(path: string): Promise<string> {
    const response = await fetch(path);
    if (!response.ok) {
      throw new Error(`Failed to read file: ${response.statusText}`);
    }
    return await response.text();
  }

  async writeFile(path: string, data: Uint8Array | string): Promise<void> {
    // Web 环境无法直接写入文件，需要用户交互
    const blob =
      data instanceof Uint8Array
        ? new Blob([data.buffer as any])
        : new Blob([data], { type: 'text/plain' });

    // 触发下载
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = this.basename(path);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  async deleteFile(path: string): Promise<void> {
    throw new Error('File deletion is not supported in web environment');
  }

  async exists(path: string): Promise<boolean> {
    try {
      const response = await fetch(path, { method: 'HEAD' });
      return response.ok;
    } catch {
      return false;
    }
  }

  // 目录操作 - Web 实现
  async listDirectory(path: string): Promise<FileInfo[]> {
    // Web 无法直接列出目录，返回空数组或使用虚拟目录结构
    return [];
  }

  async createDirectory(path: string): Promise<void> {
    throw new Error('Directory creation is not supported in web environment');
  }

  async deleteDirectory(path: string): Promise<void> {
    throw new Error('Directory deletion is not supported in web environment');
  }

  // 路径操作
  joinPath(...paths: string[]): string {
    return paths.join('/').replace(/\/+/g, '/');
  }

  dirname(path: string): string {
    const parts = path.split('/');
    parts.pop();
    return parts.join('/') || '/';
  }

  basename(path: string): string {
    const parts = path.split('/');
    return parts[parts.length - 1] || '';
  }

  normalize(path: string): string {
    return path.replace(/\/+/g, '/');
  }

  async joinPathAsync(...paths: string[]): Promise<string> {
    return this.joinPath(...paths);
  }

  async dirnameAsync(path: string): Promise<string> {
    return this.dirname(path);
  }

  async basenameAsync(path: string): Promise<string> {
    return this.basename(path);
  }

  async normalizeAsync(path: string): Promise<string> {
    return this.normalize(path);
  }

  // Office 文件特定操作 - Web 实现
  async parseDocx(path: string): Promise<DocxContent> {
    // 在 Web 环境中，需要读取文件并使用 JS 库解析
    const data = await this.readFileAsText(path);
    return {
      text: data,
      paragraphs: [
        {
          id: '1',
          text: data,
          runs: [
            {
              text: data,
            },
          ],
        },
      ],
      tables: [],
      metadata: {},
    };
  }

  async parseXlsx(path: string): Promise<XlsxContent> {
    // 在 Web 环境中，需要使用 xlsx 库等 JS 库解析
    const data = await this.readFileAsText(path);
    return {
      sheets: [
        {
          id: 'sheet1',
          name: 'Sheet1',
          cells: new Map(),
        },
      ],
      metadata: {},
    };
  }

  async saveDocx(path: string, content: DocxContent): Promise<void> {
    // Web 环境需要生成并下载 docx
    const text = content.paragraphs.map(p => p.text).join('\n');
    await this.writeFile(path, text);
  }

  async saveXlsx(path: string, content: XlsxContent): Promise<void> {
    // Web 环境需要生成并下载 xlsx
    await this.writeFile(path, JSON.stringify(content));
  }

  // 文件监听 - Web 实现 (轮询方式)
  watchFile(path: string, callback: FileWatcherCallback): UnwatchFn {
    if (!this.watchers.has(path)) {
      this.watchers.set(path, new Set());
    }
    this.watchers.get(path)!.add(callback);

    // Web 环境使用轮询检测文件变化
    let lastModified = 0;
    const checkInterval = setInterval(async () => {
      try {
        const response = await fetch(path, { method: 'HEAD' });
        const modified = Number(response.headers.get('Last-Modified') || 0);
        if (lastModified > 0 && modified !== lastModified) {
          callback({
            type: 'modified',
            path,
            timestamp: Date.now(),
          });
        }
        lastModified = modified;
      } catch {
        // 忽略错误
      }
    }, 5000); // 5秒轮询一次

    return () => {
      this.watchers.get(path)?.delete(callback);
      clearInterval(checkInterval);
    };
  }

  watchDirectory(path: string, callback: FileWatcherCallback): UnwatchFn {
    // Web 环境不支持目录监听
    return () => {};
  }

  // 桌面特有功能 - Web 实现
  async selectFile(filters?: FileFilter[]): Promise<string | null> {
    const files = await this.openFileDialog({
      filters: filters || [FILE_FILTERS.all],
      multiple: false,
    });
    return files[0] || null;
  }

  async selectDirectory(): Promise<string | null> {
    throw new Error('Directory selection is not supported in web environment');
  }

  async showInFolder(path: string): Promise<void> {
    throw new Error('Show in folder is not supported in web environment');
  }

  async openFileDialog(options: OpenFileDialogOptions): Promise<string[]> {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.multiple = options.multiple || false;

      if (options.filters && options.filters.length > 0) {
        const extensions = options.filters.flatMap(f => f.extensions);
        input.accept = extensions.map(ext => `.${ext}`).join(',');
      }

      input.onchange = (e) => {
        const target = e.target as HTMLInputElement;
        const files = Array.from(target.files || []);
        // Web 环境返回临时 URL
        resolve(files.map(f => URL.createObjectURL(f)));
      };

      input.oncancel = () => resolve([]);
      input.click();
    });
  }

  async saveFileDialog(options: SaveFileDialogOptions): Promise<string | null> {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.setAttribute('nwworkingdir', options.defaultPath || '');

      input.onchange = (e) => {
        const target = e.target as HTMLInputElement;
        resolve(target.files?.[0]?.name || null);
      };

      input.oncancel = () => resolve(null);
      input.click();
    });
  }

  // 获取系统路径 - Web 实现
  async getAppDataPath(): Promise<string> {
    return '/app-data';
  }

  async getDocumentsPath(): Promise<string> {
    return '/documents';
  }

  async getDownloadsPath(): Promise<string> {
    return '/downloads';
  }

  async getHomePath(): Promise<string> {
    return '/home';
  }
}

// Export singleton instance
export const webFileAdapter = new WebFileAdapter();
