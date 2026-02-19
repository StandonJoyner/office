// Adapter Factory and Exports
export * from './filesystem';
export * from './system';
export * from './storage';

// Web adapters
export { webFileAdapter } from './web/FileAdapterWeb';
export { webSystemAdapter } from './web/SystemAdapterWeb';

// Tauri adapters
export {
  tauriFileAdapter,
  createTauriFileAdapter,
  isTauriEnvironment as isTauriFileEnvironment,
} from './tauri/FileAdapterTauri';
export {
  tauriSystemAdapter,
  createTauriSystemAdapter,
} from './tauri/SystemAdapterTauri';
export {
  tauriStorageAdapter,
  createTauriStorageAdapter,
} from './tauri/storage';

// Re-export platform helpers without conflict
export { isDesktop, getPlatform } from './filesystem';

import type { FileAdapter } from './filesystem';
import type { SystemAdapter } from './system';
import { isDesktop } from './filesystem';
import { isTauri as isTauriSystem } from './system';
import {
  createTauriFileAdapter,
  tauriFileAdapter,
  isTauriEnvironment as isTauriFileEnvironment,
} from './tauri/FileAdapterTauri';
import { webFileAdapter } from './web/FileAdapterWeb';
import { webSystemAdapter } from './web/SystemAdapterWeb';
import {
  createTauriSystemAdapter,
  tauriSystemAdapter,
} from './tauri/SystemAdapterTauri';

/**
 * File Adapter Factory
 * 根据运行环境返回适当的文件系统适配器
 */
export async function createFileAdapter(): Promise<FileAdapter> {
  // 检查是否在 Tauri 环境中
  if (isTauriSystem() || isTauriFileEnvironment()) {
    return await createTauriFileAdapter();
  }

  // 检查是否在其他桌面环境 (Electron 等)
  if (isDesktop()) {
    // TODO: 实现 Electron 适配器
    console.warn('Desktop environment detected but no adapter available, using web adapter');
    return webFileAdapter;
  }

  // 默认使用 Web 适配器
  return webFileAdapter;
}

/**
 * 获取文件系统适配器单例
 * 首次调用会根据环境创建实例
 */
let fileAdapterInstance: FileAdapter | null = null;

export async function getFileAdapter(): Promise<FileAdapter> {
  if (!fileAdapterInstance) {
    fileAdapterInstance = await createFileAdapter();
  }
  return fileAdapterInstance;
}

/**
 * 重置文件系统适配器（主要用于测试）
 */
export function resetFileAdapter(): void {
  fileAdapterInstance = null;
}

/**
 * 导出便捷函数
 */
export const fileSystem = {
  readFile: async (path: string) => (await getFileAdapter()).readFile(path),
  readFileAsText: async (path: string) => (await getFileAdapter()).readFileAsText(path),
  writeFile: async (path: string, data: Uint8Array | string) =>
    (await getFileAdapter()).writeFile(path, data),
  deleteFile: async (path: string) => (await getFileAdapter()).deleteFile(path),
  exists: async (path: string) => (await getFileAdapter()).exists(path),
  listDirectory: async (path: string) => (await getFileAdapter()).listDirectory(path),
  createDirectory: async (path: string) => (await getFileAdapter()).createDirectory(path),
  deleteDirectory: async (path: string) => (await getFileAdapter()).deleteDirectory(path),
  joinPath: (...paths: string[]) => getFileAdapter().then(a => a.joinPath(...paths)),
  dirname: (path: string) => getFileAdapter().then(a => a.dirname(path)),
  basename: (path: string) => getFileAdapter().then(a => a.basename(path)),
  parseDocx: (path: string) => getFileAdapter().then(a => a.parseDocx(path)),
  parseXlsx: (path: string) => getFileAdapter().then(a => a.parseXlsx(path)),
  saveDocx: (path: string, content: any) => getFileAdapter().then(a => a.saveDocx(path, content)),
  saveXlsx: (path: string, content: any) => getFileAdapter().then(a => a.saveXlsx(path, content)),
  watchFile: (path: string, callback: any) => getFileAdapter().then(a => a.watchFile(path, callback)),
  selectFile: (filters?: any[]) => getFileAdapter().then(a => a.selectFile(filters)),
  selectDirectory: () => getFileAdapter().then(a => a.selectDirectory()),
  showInFolder: (path: string) => getFileAdapter().then(a => a.showInFolder(path)),
  openFileDialog: (options: any) => getFileAdapter().then(a => a.openFileDialog(options)),
  saveFileDialog: (options: any) => getFileAdapter().then(a => a.saveFileDialog(options)),
  getAppDataPath: async () => (await getFileAdapter()).getAppDataPath(),
  getDocumentsPath: async () => (await getFileAdapter()).getDocumentsPath(),
  getDownloadsPath: async () => (await getFileAdapter()).getDownloadsPath(),
  getHomePath: async () => (await getFileAdapter()).getHomePath(),
};

/**
 * System Adapter Factory
 * 根据运行环境返回适当的系统适配器
 */
export async function createSystemAdapter(): Promise<SystemAdapter> {
  // 检查是否在 Tauri 环境中
  if (isTauriSystem()) {
    return await createTauriSystemAdapter();
  }

  // 检查是否在其他桌面环境 (Electron 等)
  if (isDesktop()) {
    // TODO: 实现 Electron 适配器
    console.warn('Desktop environment detected but no adapter available, using web adapter');
    return webSystemAdapter;
  }

  // 默认使用 Web 适配器
  return webSystemAdapter;
}

/**
 * 获取系统适配器单例
 * 首次调用会根据环境创建实例
 */
let systemAdapterInstance: SystemAdapter | null = null;

export async function getSystemAdapter(): Promise<SystemAdapter> {
  if (!systemAdapterInstance) {
    systemAdapterInstance = await createSystemAdapter();
  }
  return systemAdapterInstance;
}

/**
 * 重置系统适配器（主要用于测试）
 */
export function resetSystemAdapter(): void {
  systemAdapterInstance = null;
}

/**
 * 导出便捷函数
 */
export const system = {
  minimizeWindow: async () => (await getSystemAdapter()).minimizeWindow(),
  maximizeWindow: async () => (await getSystemAdapter()).maximizeWindow(),
  restoreWindow: async () => (await getSystemAdapter()).restoreWindow(),
  closeWindow: async () => (await getSystemAdapter()).closeWindow(),
  setWindowTitle: async (title: string) => (await getSystemAdapter()).setWindowTitle(title),
  enterFullscreen: async () => (await getSystemAdapter()).enterFullscreen(),
  exitFullscreen: async () => (await getSystemAdapter()).exitFullscreen(),
  toggleFullscreen: async () => (await getSystemAdapter()).toggleFullscreen(),
  createMenu: async (items: any[]) => (await getSystemAdapter()).createMenu(items),
  showNotification: async (options: any) => (await getSystemAdapter()).showNotification(options),
  readText: async () => (await getSystemAdapter()).readText(),
  writeText: async (text: string) => (await getSystemAdapter()).writeText(text),
  alert: async (message: string) => (await getSystemAdapter()).alert(message),
  confirm: async (message: string) => (await getSystemAdapter()).confirm(message),
  getVersion: async () => (await getSystemAdapter()).getVersion(),
  openFile: async (path: string) => (await getSystemAdapter()).openFile(path),
  openUrl: async (url: string) => (await getSystemAdapter()).openUrl(url),
  enableDragDrop: async (callback: any) => (await getSystemAdapter()).enableDragDrop(callback),
  disableDragDrop: async () => (await getSystemAdapter()).disableDragDrop(),
  getSystemTheme: async () => (await getSystemAdapter()).getSystemTheme(),
};
