// Tauri System Adapter Implementation
import type {
  SystemAdapter,
  MenuItem,
  NotificationOptions,
  DragDropCallback,
  NotificationPermission,
} from '../system';

/**
 * Tauri System Adapter
 * 使用 Tauri 的窗口、菜单、通知等 API
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

// Type helpers for Tauri modules
async function getWindow() {
  // @ts-ignore - Tauri API
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  return getCurrentWindow();
}

async function getApp() {
  // @ts-ignore - Tauri API
  const { getVersion, getName } = await import('@tauri-apps/api/app');
  return { getVersion, getName };
}

async function getDialog() {
  // @ts-ignore - Tauri API
  const { ask, message, open, save } = await import('@tauri-apps/plugin-dialog');
  return { ask, message, open, save };
}

async function getShell() {
  // @ts-ignore - Tauri API
  const { open } = await import('@tauri-apps/plugin-opener');
  return { open };
}

class TauriSystemAdapter implements SystemAdapter {
  private dragDropCallback: DragDropCallback | null = null;
  private themeUnlisten: (() => void) | null = null;

  // 窗口管理 - Tauri 实现
  async minimizeWindow(): Promise<void> {
    const window = await getWindow();
    await window.minimize();
  }

  async maximizeWindow(): Promise<void> {
    const window = await getWindow();
    await window.maximize();
  }

  async restoreWindow(): Promise<void> {
    const window = await getWindow();
    await window.unmaximize();
  }

  async closeWindow(): Promise<void> {
    const window = await getWindow();
    await window.close();
  }

  async setWindowTitle(title: string): Promise<void> {
    const window = await getWindow();
    await window.setTitle(title);
  }

  async setWindowResizable(resizable: boolean): Promise<void> {
    const window = await getWindow();
    await window.setResizable(resizable);
  }

  async setWindowSize(width: number, height: number): Promise<void> {
    const window = await getWindow();
    // Import PhysicalSize from Tauri window API
    // @ts-ignore - Tauri API
    const { LogicalSize } = await import('@tauri-apps/api/dpi');
    await window.setSize(new LogicalSize(width, height));
  }

  // 全屏控制 - Tauri 实现
  async enterFullscreen(): Promise<void> {
    const window = await getWindow();
    await window.setFullscreen(true);
  }

  async exitFullscreen(): Promise<void> {
    const window = await getWindow();
    await window.setFullscreen(false);
  }

  async toggleFullscreen(): Promise<void> {
    const window = await getWindow();
    const isFullscreen = await window.isFullscreen();
    await window.setFullscreen(!isFullscreen);
  }

  // 系统菜单 - Tauri 实现
  async showMenu(): Promise<void> {
    // Tauri 菜单始终显示，这里可以用于显示上下文菜单
    await invoke('show_menu');
  }

  async hideMenu(): Promise<void> {
    // Tauri 不支持隐藏菜单
    console.warn('Hide menu not supported in Tauri');
  }

  async createMenu(menuItems: MenuItem[]): Promise<void> {
    // 使用 Tauri 的 menu API
    await invoke('create_menu', { items: this.convertMenuItems(menuItems) });
  }

  private convertMenuItems(items: MenuItem[]): any[] {
    return items.map(item => ({
      id: item.id,
      label: item.label,
      type: item.type,
      enabled: item.enabled !== false,
      checked: item.checked,
      accelerator: item.accelerator,
      submenu: item.submenu ? this.convertMenuItems(item.submenu) : undefined,
    }));
  }

  // 通知 - Tauri 实现
  async showNotification(options: NotificationOptions): Promise<void> {
    try {
      // 使用 Tauri 的 notification API
      // @ts-ignore - Tauri API
      const { sendNotification } = await import('@tauri-apps/plugin-notification');

      const notificationId = await sendNotification({
        title: options.title,
        body: options.body,
        icon: options.icon,
        sound: options.silent ? 'default' : undefined,
      });

      // 监听通知点击（需要 Rust 端支持）
      if (options.onClick || options.onClose) {
        invoke('listen_notification', {
          id: notificationId,
          onClick: options.onClick,
          onClose: options.onClose,
        }).catch(e => {
          console.warn('Failed to listen to notification:', e);
        });
      }
    } catch (e: any) {
      // 如果 Tauri notification 插件不可用，使用 Web API
      console.warn('Tauri notification not available, falling back to web:', e);
      this.showWebNotification(options);
    }
  }

  private async showWebNotification(options: NotificationOptions): Promise<void> {
    const permission = await this.requestNotificationPermission();
    if (permission !== 'granted') {
      console.warn('Notification permission denied');
      return;
    }

    const notification = new Notification(options.title, {
      body: options.body,
      icon: options.icon,
      silent: options.silent,
    });

    notification.onclick = () => {
      options.onClick?.();
      notification.close();
    };

    notification.onclose = () => {
      options.onClose?.();
    };
  }

  async requestNotificationPermission(): Promise<NotificationPermission> {
    try {
      // Tauri notification 权限通常在安装时授予
      // @ts-ignore - Tauri API
      const { isPermissionGranted, requestPermission } = await import('@tauri-apps/plugin-notification');

      const granted = await isPermissionGranted();
      if (granted) {
        return 'granted';
      }

      const permission = await requestPermission();
      return permission === 'granted' ? 'granted' : 'denied';
    } catch {
      // 回退到 Web API
      if (!('Notification' in window)) {
        return 'denied';
      }
      if (Notification.permission === 'granted') {
        return 'granted';
      }
      if (Notification.permission !== 'denied') {
        const permission = await Notification.requestPermission();
        return permission as NotificationPermission;
      }
      return Notification.permission as NotificationPermission;
    }
  }

  // 剪贴板 - Tauri 实现
  async readText(): Promise<string> {
    try {
      // @ts-ignore - Tauri API
      const { readText } = await import('@tauri-apps/plugin-clipboard-manager');
      return await readText();
    } catch (e: any) {
      console.warn('Tauri clipboard not available, falling back to web:', e);
      return await navigator.clipboard.readText();
    }
  }

  async writeText(text: string): Promise<void> {
    try {
      // @ts-ignore - Tauri API
      const { writeText } = await import('@tauri-apps/plugin-clipboard-manager');
      await writeText(text);
    } catch (e: any) {
      console.warn('Tauri clipboard not available, falling back to web:', e);
      await navigator.clipboard.writeText(text);
    }
  }

  async readImage(): Promise<Uint8Array> {
    try {
      // @ts-ignore - Tauri API
      const { readImage } = await import('@tauri-apps/plugin-clipboard-manager');
      const image = await readImage();
      // @ts-ignore - Tauri API
      return image.toBuffer();
    } catch (e: any) {
      throw new Error(`Failed to read image from clipboard: ${e.message}`);
    }
  }

  async writeImage(data: Uint8Array): Promise<void> {
    try {
      // @ts-ignore - Tauri API
      const { writeImage } = await import('@tauri-apps/plugin-clipboard-manager');
      await writeImage(data);
    } catch (e: any) {
      throw new Error(`Failed to write image to clipboard: ${e.message}`);
    }
  }

  // 系统对话框 - Tauri 实现
  async alert(message: string): Promise<void> {
    const dialog = await getDialog();
    await dialog.message(message, { title: 'Office Suite' });
  }

  async confirm(message: string): Promise<boolean> {
    const dialog = await getDialog();
    const confirmed = await dialog.ask(message, {
      title: 'Office Suite',
      kind: 'warning',
    });
    return confirmed === true;
  }

  async prompt(message: string, defaultText?: string): Promise<string | null> {
    // Tauri dialog 没有原生 prompt，使用自定义对话框
    const result = await invoke<string | null>('prompt', {
      message,
      defaultText: defaultText || '',
    });
    return result;
  }

  // 应用信息 - Tauri 实现
  async getVersion(): Promise<string> {
    const app = await getApp();
    return await app.getVersion();
  }

  async getName(): Promise<string> {
    const app = await getApp();
    return await app.getName();
  }

  // 文件关联 - Tauri 实现
  async openFile(path: string): Promise<void> {
    const shell = await getShell();
    await shell.open(path);
  }

  async openUrl(url: string): Promise<void> {
    const shell = await getShell();
    await shell.open(url);
  }

  // 拖拽支持 - Tauri 实现
  enableDragDrop(callback: DragDropCallback): void {
    this.dragDropCallback = callback;
    // Tauri 的文件拖拽需要在 Rust 端配置
    // 这里注册前端事件监听
    invoke('enable_drag_drop').catch(e => {
      console.warn('Failed to enable drag drop:', e);
    });

    // 监听来自 Rust 的拖拽事件
    window.addEventListener('tauri-drag-drop', this.handleTauriDragDrop as EventListener);
  }

  disableDragDrop(): void {
    this.dragDropCallback = null;
    invoke('disable_drag_drop').catch(e => {
      console.warn('Failed to disable drag drop:', e);
    });
    window.removeEventListener('tauri-drag-drop', this.handleTauriDragDrop as EventListener);
  }

  private handleTauriDragDrop = (e: any): void => {
    const { type, files, text } = e.detail;

    if (type === 'drag-enter' && this.dragDropCallback?.onDragEnter) {
      this.dragDropCallback.onDragEnter();
    } else if (type === 'drag-leave' && this.dragDropCallback?.onDragLeave) {
      this.dragDropCallback.onDragLeave();
    } else if (type === 'drop') {
      if (files && files.length > 0 && this.dragDropCallback?.onFileDrop) {
        this.dragDropCallback.onFileDrop(files);
      }
      if (text && this.dragDropCallback?.onTextDrop) {
        this.dragDropCallback.onTextDrop(text);
      }
    }
  };

  // 系统主题 - Tauri 实现
  async getSystemTheme(): Promise<'light' | 'dark' | 'auto'> {
    try {
      const theme = await invoke<'light' | 'dark'>('get_theme');
      return theme;
    } catch (e) {
      // 回退到 Web API
      if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        return 'dark';
      }
      return 'light';
    }
  }

  onThemeChange(callback: (theme: 'light' | 'dark') => void): () => void {
    // 使用 Tauri 的事件系统监听主题变化
    invoke('listen_theme_change')
      .then((unlisten: any) => {
        this.themeUnlisten = unlisten;
      })
      .catch(e => {
        console.warn('Failed to listen to theme change:', e);
      });

    // 同时设置前端监听作为回退
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');

    const handler = (e: MediaQueryListEvent) => {
      callback(e.matches ? 'dark' : 'light');
    };

    mediaQuery.addEventListener('change', handler);

    return () => {
      this.themeUnlisten?.();
      mediaQuery.removeEventListener('change', handler);
    };
  }

  // 额外功能：窗口位置
  async setWindowPosition(x: number, y: number): Promise<void> {
    const window = await getWindow();
    // Import LogicalPosition from Tauri window API
    // @ts-ignore - Tauri API
    const { LogicalPosition } = await import('@tauri-apps/api/dpi');
    await window.setPosition(new LogicalPosition(x, y));
  }

  async getWindowPosition(): Promise<{ x: number; y: number }> {
    const window = await getWindow();
    return await window.outerPosition();
  }

  async setDecorations(hasDecorations: boolean): Promise<void> {
    const window = await getWindow();
    await window.setDecorations(hasDecorations);
  }

  async setAlwaysOnTop(alwaysOnTop: boolean): Promise<void> {
    const window = await getWindow();
    await window.setAlwaysOnTop(alwaysOnTop);
  }

  async isAlwaysOnTop(): Promise<boolean> {
    const window = await getWindow();
    return await window.isAlwaysOnTop();
  }
}

// 导出 Tauri 系统适配器实例
export let tauriSystemAdapter: TauriSystemAdapter | null = null;

/**
 * 创建 Tauri 系统适配器实例
 */
export async function createTauriSystemAdapter(): Promise<TauriSystemAdapter> {
  if (!tauriSystemAdapter) {
    tauriSystemAdapter = new TauriSystemAdapter();
  }
  return tauriSystemAdapter;
}
