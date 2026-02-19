// Tauri Storage Adapter Implementation
import type { StorageAdapter } from '../storage';

/**
 * Tauri Storage Adapter
 * 使用 Tauri 的文件系统和路径 API 实现持久化存储
 */

// Tauri invoke helper
async function invoke<T>(command: string, args?: any): Promise<T> {
  try {
    // @ts-ignore - Tauri global
    const { invoke } = await import('@tauri-apps/api/core');
    return await invoke(command, args);
  } catch (e: any) {
    throw new Error(`Tauri invoke failed (${command}): ${e.message}`);
  }
}

class TauriStorageAdapter implements StorageAdapter {
  private storagePath: string | null = null;
  private initialized = false;

  async init(): Promise<void> {
    if (this.initialized) return;

    try {
      // 获取应用数据目录
      // @ts-ignore - Tauri API
      const { appDataDir } = await import('@tauri-apps/api/path');
      const appPath = await appDataDir();
      this.storagePath = appPath || '';

      // 确保存储目录存在
      if (!this.storagePath) {
        throw new Error('Storage path initialization failed');
      }
      // @ts-ignore
      const dir = await this.joinPath(this.storagePath, 'office-suite', 'storage');
      await invoke('create_dir', { path: (dir as string) || '', recursive: true });

      this.initialized = true;
    } catch (e: any) {
      console.error('Failed to initialize Tauri storage:', e);
      throw e;
    }
  }

  private async ensureInitialized(): Promise<void> {
    if (!this.initialized) {
      await this.init();
    }
  }

  private async getStorageDir(): Promise<string> {
    await this.ensureInitialized();
    if (!this.storagePath) {
      throw new Error('Storage path not initialized');
    }
    const path = await this.joinPath(this.storagePath, 'office-suite', 'storage');
    return path || '';
  }

  private async joinPath(...paths: string[]): Promise<string> {
    // @ts-ignore - Tauri API
    const { join } = await import('@tauri-apps/api/path');
    return await join(...paths);
  }

  private sanitizeKey(key: string): string {
    // 替换不安全的文件名字符
    return key.replace(/[^a-zA-Z0-9_-]/g, '_');
  }

  private async getKeyPath(key: string): Promise<string> {
    const dir = await this.getStorageDir();
    const safeKey = this.sanitizeKey(key);
    const result = await this.joinPath(dir, `${safeKey}.json`);
    return result || '';
  }

  async save(key: string, value: any): Promise<void> {
    await this.ensureInitialized();
    const keyPath = await this.getKeyPath(key);

    try {
      const data = JSON.stringify({ key, value, timestamp: Date.now() });
      await invoke('write_text_file', { path: keyPath, contents: data });
    } catch (e: any) {
      throw new Error(`Failed to save key "${key}": ${e.message}`);
    }
  }

  async load<T = any>(key: string): Promise<T | null> {
    await this.ensureInitialized();
    const keyPath = await this.getKeyPath(key);

    try {
      // 检查文件是否存在
      const exists = await invoke<boolean>('exists', { path: keyPath });
      if (!exists) {
        return null;
      }

      const content = await invoke<string>('read_text_file', { path: keyPath });
      const parsed = JSON.parse(content);
      return parsed.value as T;
    } catch (e: any) {
      // 如果是文件不存在错误，返回 null
      if (e.message?.includes('No such file') || e.message?.includes('not found')) {
        return null;
      }
      throw new Error(`Failed to load key "${key}": ${e.message}`);
    }
  }

  async delete(key: string): Promise<void> {
    await this.ensureInitialized();
    const keyPath = await this.getKeyPath(key);

    try {
      const exists = await invoke<boolean>('exists', { path: keyPath });
      if (exists) {
        await invoke('remove_file', { path: keyPath });
      }
    } catch (e: any) {
      throw new Error(`Failed to delete key "${key}": ${e.message}`);
    }
  }

  async exists(key: string): Promise<boolean> {
    await this.ensureInitialized();
    const keyPath = await this.getKeyPath(key);
    return await invoke<boolean>('exists', { path: keyPath });
  }

  async clear(): Promise<void> {
    await this.ensureInitialized();
    const storageDir = await this.getStorageDir();

    try {
      // 删除整个存储目录
      const exists = await invoke<boolean>('exists', { path: storageDir });
      if (exists) {
        await invoke('remove_dir', { path: storageDir, recursive: true });
        // 重新创建目录
        await invoke('create_dir', { path: storageDir, recursive: true });
      }
    } catch (e: any) {
      throw new Error(`Failed to clear storage: ${e.message}`);
    }
  }

  async getAllKeys(): Promise<string[]> {
    await this.ensureInitialized();
    const storageDir = await this.getStorageDir();

    try {
      const entries = await invoke<any[]>('read_dir', { path: storageDir, recursive: false });
      return entries
        .filter((entry: any) => entry.metadata.is_file && entry.path.endsWith('.json'))
        .map((entry: any) => {
          const filename = entry.path.split(/[/\\]/).pop() || '';
          return filename.replace('.json', '');
        });
    } catch (e: any) {
      throw new Error(`Failed to get all keys: ${e.message}`);
    }
  }

  async getSize(): Promise<number> {
    await this.ensureInitialized();
    const storageDir = await this.getStorageDir();

    try {
      const entries = await invoke<any[]>('read_dir', { path: storageDir, recursive: false });
      return entries
        .filter((entry: any) => entry.metadata.is_file)
        .reduce((total: number, entry: any) => total + entry.metadata.size, 0);
    } catch (e: any) {
      throw new Error(`Failed to get storage size: ${e.message}`);
    }
  }
}

// 导出 Tauri 存储适配器实例
export let tauriStorageAdapter: TauriStorageAdapter | null = null;

/**
 * 创建 Tauri 存储适配器实例
 */
export async function createTauriStorageAdapter(): Promise<TauriStorageAdapter> {
  if (!tauriStorageAdapter) {
    tauriStorageAdapter = new TauriStorageAdapter();
    await tauriStorageAdapter.init();
  }
  return tauriStorageAdapter;
}
