// Storage Adapter Interface for Web/Desktop abstraction
export interface StorageAdapter {
  save(key: string, value: any): Promise<void>;
  load<T = any>(key: string): Promise<T | null>;
  delete(key: string): Promise<void>;
  exists(key: string): Promise<boolean>;
  clear(): Promise<void>;
}

// Web Implementation using IndexedDB
class BrowserStorageAdapter implements StorageAdapter {
  private db: IDBDatabase | null = null;
  private readonly DB_NAME = 'OfficeSuiteDB';
  private readonly DB_VERSION = 1;
  private readonly STORE_NAME = 'documents';

  async init(): Promise<void> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.DB_NAME, this.DB_VERSION);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        this.db = request.result;
        resolve();
      };

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(this.STORE_NAME)) {
          const store = db.createObjectStore(this.STORE_NAME, { keyPath: 'key' });
          store.createIndex('key', 'key', { unique: true });
        }
      };
    });
  }

  private ensureInitialized(): void {
    if (!this.db) {
      throw new Error('Storage not initialized. Call init() first.');
    }
  }

  async save(key: string, value: any): Promise<void> {
    this.ensureInitialized();
    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.STORE_NAME], 'readwrite');
      const store = transaction.objectStore(this.STORE_NAME);
      const request = store.put({ key, value, timestamp: Date.now() });

      transaction.onerror = () => reject(transaction.error);
      transaction.oncomplete = () => resolve();
    });
  }

  async load<T = any>(key: string): Promise<T | null> {
    this.ensureInitialized();
    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.STORE_NAME], 'readonly');
      const store = transaction.objectStore(this.STORE_NAME);
      const request = store.get(key);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        const result = request.result;
        resolve(result ? result.value : null);
      };
    });
  }

  async delete(key: string): Promise<void> {
    this.ensureInitialized();
    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.STORE_NAME], 'readwrite');
      const store = transaction.objectStore(this.STORE_NAME);
      const request = store.delete(key);

      transaction.onerror = () => reject(transaction.error);
      transaction.oncomplete = () => resolve();
    });
  }

  async exists(key: string): Promise<boolean> {
    this.ensureInitialized();
    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.STORE_NAME], 'readonly');
      const store = transaction.objectStore(this.STORE_NAME);
      const request = store.getKey(key);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve(request.result !== undefined);
    });
  }

  async clear(): Promise<void> {
    this.ensureInitialized();
    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.STORE_NAME], 'readwrite');
      const store = transaction.objectStore(this.STORE_NAME);
      const request = store.clear();

      transaction.onerror = () => reject(transaction.error);
      transaction.oncomplete = () => resolve();
    });
  }
}

// Platform detection helpers
function isTauriEnvironment(): boolean {
  return typeof window !== 'undefined' && (window as any).__TAURI__ !== undefined;
}

function isElectronEnvironment(): boolean {
  return typeof window !== 'undefined' && typeof (window as any).process === 'object';
}

/**
 * Factory function to get the appropriate storage adapter
 */
export function createStorageAdapter(): StorageAdapter {
  const adapter = new BrowserStorageAdapter();
  adapter.init();
  return adapter;
}

/**
 * Async factory function to get the appropriate storage adapter
 * 支持桌面环境
 */
export async function createStorageAdapterAsync(): Promise<StorageAdapter> {
  if (isTauriEnvironment()) {
    const { createTauriStorageAdapter } = await import('./tauri/storage');
    return await createTauriStorageAdapter();
  }

  // TODO: 添加 Electron 支持
  if (isElectronEnvironment()) {
    throw new Error('Electron storage adapter not implemented yet');
  }

  const adapter = new BrowserStorageAdapter();
  await adapter.init();
  return adapter;
}

// 同步获取（兼容现有代码）
export const storage = createStorageAdapter();
