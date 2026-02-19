// System Adapter Interface for Web/Desktop abstraction

export interface SystemAdapter {
  // 窗口管理
  minimizeWindow(): Promise<void>;
  maximizeWindow(): Promise<void>;
  restoreWindow(): Promise<void>;
  closeWindow(): Promise<void>;
  setWindowTitle(title: string): Promise<void>;
  setWindowResizable(resizable: boolean): Promise<void>;
  setWindowSize(width: number, height: number): Promise<void>;

  // 全屏控制
  enterFullscreen(): Promise<void>;
  exitFullscreen(): Promise<void>;
  toggleFullscreen(): Promise<void>;

  // 系统菜单
  showMenu(): Promise<void>;
  hideMenu(): Promise<void>;
  createMenu(menuItems: MenuItem[]): Promise<void>;

  // 通知
  showNotification(options: NotificationOptions): Promise<void>;
  requestNotificationPermission(): Promise<NotificationPermission>;

  // 剪贴板
  readText(): Promise<string>;
  writeText(text: string): Promise<void>;
  readImage(): Promise<Uint8Array>;
  writeImage(data: Uint8Array): Promise<void>;

  // 系统对话框
  alert(message: string): Promise<void>;
  confirm(message: string): Promise<boolean>;
  prompt(message: string, defaultText?: string): Promise<string | null>;

  // 应用信息
  getVersion(): Promise<string>;
  getName(): Promise<string>;

  // 文件关联
  openFile(path: string): Promise<void>;
  openUrl(url: string): Promise<void>;

  // 拖拽支持
  enableDragDrop(callback: DragDropCallback): void;
  disableDragDrop(): void;

  // 系统主题
  getSystemTheme(): Promise<'light' | 'dark' | 'auto'>;
  onThemeChange(callback: (theme: 'light' | 'dark') => void): () => void;
}

// Type Definitions
export interface MenuItem {
  id?: string;
  label: string;
  type?: 'normal' | 'separator' | 'submenu' | 'checkbox' | 'radio';
  enabled?: boolean;
  checked?: boolean;
  accelerator?: string;
  click?: () => void | Promise<void>;
  submenu?: MenuItem[];
}

export interface NotificationOptions {
  title: string;
  body: string;
  icon?: string;
  silent?: boolean;
  onClick?: () => void;
  onClose?: () => void;
}

export type NotificationPermission = 'granted' | 'denied' | 'default';

export interface DragDropCallback {
  onFileDrop?: (files: string[]) => void;
  onTextDrop?: (text: string) => void;
  onDragEnter?: () => void;
  onDragLeave?: () => void;
}

// Platform Detection
export function isDesktop(): boolean {
  if (typeof window === 'undefined') return false;
  // Check for Tauri
  if ((window as any).__TAURI__) return true;
  // Check for Electron
  if (typeof (window as any).process === 'object' && (window as any).process.versions?.electron) return true;
  return false;
}

export function isTauri(): boolean {
  return typeof window !== 'undefined' && (window as any).__TAURI__ !== undefined;
}

export function isElectron(): boolean {
  return (
    typeof window !== 'undefined' &&
    typeof (window as any).process === 'object' &&
    (window as any).process.versions?.electron
  );
}

export function getPlatform(): 'windows' | 'macos' | 'linux' | 'web' {
  if (isTauri()) {
    return (window as any).__TAURI__.os.platform() || 'web';
  }
  if (isElectron()) {
    const platform = (window as any).process.platform;
    if (platform === 'win32') return 'windows';
    if (platform === 'darwin') return 'macos';
    if (platform === 'linux') return 'linux';
  }
  if (navigator.platform.startsWith('Win')) return 'windows';
  if (navigator.platform.startsWith('Mac')) return 'macos';
  if (navigator.platform.startsWith('Linux')) return 'linux';
  return 'web';
}

// Menu presets
export const MENU_PRESETS = {
  // macOS 应用菜单
  macos: [
    {
      label: 'Office Suite',
      submenu: [
        { label: 'About', accelerator: 'CmdOrCtrl+Shift+A', click: () => console.log('About') },
        { type: 'separator' },
        { label: 'Preferences', accelerator: 'CmdOrCtrl+,', click: () => console.log('Preferences') },
        { type: 'separator' },
        { label: 'Quit', accelerator: 'CmdOrCtrl+Q', click: () => console.log('Quit') },
      ],
    },
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => console.log('New') },
        { label: 'Open', accelerator: 'CmdOrCtrl+O', click: () => console.log('Open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => console.log('Save') },
        { label: 'Save As', accelerator: 'CmdOrCtrl+Shift+S', click: () => console.log('Save As') },
        { type: 'separator' },
        { label: 'Export', submenu: [
          { label: 'Export as PDF', click: () => console.log('Export PDF') },
          { label: 'Export as Word', click: () => console.log('Export Word') },
        ]},
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: () => console.log('Undo') },
        { label: 'Redo', accelerator: 'CmdOrCtrl+Shift+Z', click: () => console.log('Redo') },
        { type: 'separator' },
        { label: 'Cut', accelerator: 'CmdOrCtrl+X', click: () => console.log('Cut') },
        { label: 'Copy', accelerator: 'CmdOrCtrl+C', click: () => console.log('Copy') },
        { label: 'Paste', accelerator: 'CmdOrCtrl+V', click: () => console.log('Paste') },
        { type: 'separator' },
        { label: 'Select All', accelerator: 'CmdOrCtrl+A', click: () => console.log('Select All') },
      ],
    },
    {
      label: 'View',
      submenu: [
        { label: 'Fullscreen', accelerator: 'F11', click: () => console.log('Fullscreen') },
        { type: 'separator' },
        { label: 'Zoom In', accelerator: 'CmdOrCtrl+Plus', click: () => console.log('Zoom In') },
        { label: 'Zoom Out', accelerator: 'CmdOrCtrl+-', click: () => console.log('Zoom Out') },
        { label: 'Reset Zoom', accelerator: 'CmdOrCtrl+0', click: () => console.log('Reset Zoom') },
      ],
    },
  ],

  // Windows/Linux 文件菜单
  fileMenu: [
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'Ctrl+N', click: () => console.log('New') },
        { label: 'Open', accelerator: 'Ctrl+O', click: () => console.log('Open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'Ctrl+S', click: () => console.log('Save') },
        { label: 'Save As', accelerator: 'Ctrl+Shift+S', click: () => console.log('Save As') },
        { type: 'separator' },
        { label: 'Exit', accelerator: 'Alt+F4', click: () => console.log('Exit') },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Undo', accelerator: 'Ctrl+Z', click: () => console.log('Undo') },
        { label: 'Redo', accelerator: 'Ctrl+Y', click: () => console.log('Redo') },
        { type: 'separator' },
        { label: 'Cut', accelerator: 'Ctrl+X', click: () => console.log('Cut') },
        { label: 'Copy', accelerator: 'Ctrl+C', click: () => console.log('Copy') },
        { label: 'Paste', accelerator: 'Ctrl+V', click: () => console.log('Paste') },
        { type: 'separator' },
        { label: 'Select All', accelerator: 'Ctrl+A', click: () => console.log('Select All') },
      ],
    },
    {
      label: 'View',
      submenu: [
        { label: 'Fullscreen', accelerator: 'F11', click: () => console.log('Fullscreen') },
        { label: 'Reload', accelerator: 'Ctrl+R', click: () => console.log('Reload') },
        { label: 'Developer Tools', accelerator: 'F12', click: () => console.log('DevTools') },
      ],
    },
  ],

  // 帮助菜单
  helpMenu: [
    {
      label: 'Help',
      submenu: [
        { label: 'Documentation', accelerator: 'F1', click: () => console.log('Documentation') },
        { label: 'Keyboard Shortcuts', accelerator: 'Ctrl+/', click: () => console.log('Shortcuts') },
        { type: 'separator' },
        { label: 'About', click: () => console.log('About') },
      ],
    },
  ],
};
