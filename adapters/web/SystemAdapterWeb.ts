// Web System Adapter Implementation
import type {
  SystemAdapter,
  MenuItem,
  NotificationOptions,
  DragDropCallback,
} from '../system';
import { MENU_PRESETS } from '../system';

class WebSystemAdapter implements SystemAdapter {
  private dragDropCallback: DragDropCallback | null = null;
  private themeCallback: ((theme: 'light' | 'dark') => void) | null = null;

  // 窗口管理 - Web 实现（有限支持）
  async minimizeWindow(): Promise<void> {
    // Web 不支持窗口最小化
    console.warn('Window minimize not supported in web environment');
  }

  async maximizeWindow(): Promise<void> {
    console.warn('Window maximize not supported in web environment');
  }

  async restoreWindow(): Promise<void> {
    console.warn('Window restore not supported in web environment');
  }

  async closeWindow(): Promise<void> {
    // 尝试关闭窗口
    if (window.opener) {
      window.close();
    } else {
      console.warn('Cannot close main window in web environment');
    }
  }

  async setWindowTitle(title: string): Promise<void> {
    document.title = title;
  }

  async setWindowResizable(resizable: boolean): Promise<void> {
    // Web 不支持
    console.warn('Window resizable not supported in web environment');
  }

  async setWindowSize(width: number, height: number): Promise<void> {
    // Web 可以调整窗口大小（在某些浏览器中）
    window.resizeTo(width, height);
  }

  // 全屏控制 - Web 实现
  async enterFullscreen(): Promise<void> {
    try {
      await document.documentElement.requestFullscreen();
    } catch (e: any) {
      console.warn('Fullscreen not supported:', e);
    }
  }

  async exitFullscreen(): Promise<void> {
    try {
      await document.exitFullscreen();
    } catch (e: any) {
      console.warn('Exit fullscreen failed:', e);
    }
  }

  async toggleFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await this.exitFullscreen();
    } else {
      await this.enterFullscreen();
    }
  }

  // 系统菜单 - Web 不支持原生菜单
  async showMenu(): Promise<void> {
    console.warn('System menu not supported in web environment');
  }

  async hideMenu(): Promise<void> {
    console.warn('System menu not supported in web environment');
  }

  async createMenu(menuItems: MenuItem[]): Promise<void> {
    console.warn('System menu not supported in web environment', menuItems);
  }

  // 通知 - Web 实现
  async showNotification(options: NotificationOptions): Promise<void> {
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

  // 剪贴板 - Web 实现
  async readText(): Promise<string> {
    return await navigator.clipboard.readText();
  }

  async writeText(text: string): Promise<void> {
    await navigator.clipboard.writeText(text);
  }

  async readImage(): Promise<Uint8Array> {
    // Web 剪贴板图像读取需要额外处理
    const items = await navigator.clipboard.read();
    for (const item of items) {
      const imageType = item.types.find(t => t.startsWith('image/'));
      if (imageType) {
        const blob = await item.getType(imageType);
        const buffer = await blob.arrayBuffer();
        return new Uint8Array(buffer);
      }
    }
    throw new Error('No image in clipboard');
  }

  async writeImage(data: Uint8Array): Promise<void> {
    const blob = new Blob([data.buffer as any], { type: 'image/png' }) as any;
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
  }

  // 系统对话框 - Web 实现
  async alert(message: string): Promise<void> {
    window.alert(message);
  }

  async confirm(message: string): Promise<boolean> {
    return window.confirm(message);
  }

  async prompt(message: string, defaultText?: string): Promise<string | null> {
    return window.prompt(message, defaultText);
  }

  // 应用信息
  async getVersion(): Promise<string> {
    return '0.1.0-web';
  }

  async getName(): Promise<string> {
    return 'Office Suite (Web)';
  }

  // 文件关联
  async openFile(path: string): Promise<void> {
    // Web 无法直接打开文件
    window.open(path, '_blank');
  }

  async openUrl(url: string): Promise<void> {
    window.open(url, '_blank');
  }

  // 拖拽支持 - Web 实现
  enableDragDrop(callback: DragDropCallback): void {
    this.dragDropCallback = callback;

    document.addEventListener('dragover', this.handleDragOver);
    document.addEventListener('dragenter', this.handleDragEnter);
    document.addEventListener('dragleave', this.handleDragLeave);
    document.addEventListener('drop', this.handleDrop);
  }

  disableDragDrop(): void {
    this.dragDropCallback = null;

    document.removeEventListener('dragover', this.handleDragOver);
    document.removeEventListener('dragenter', this.handleDragEnter);
    document.removeEventListener('dragleave', this.handleDragLeave);
    document.removeEventListener('drop', this.handleDrop);
  }

  private handleDragOver = (e: DragEvent): void => {
    e.preventDefault();
    e.dataTransfer!.dropEffect = 'copy';
  };

  private handleDragEnter = (e: DragEvent): void => {
    e.preventDefault();
    this.dragDropCallback?.onDragEnter?.();
  };

  private handleDragLeave = (e: DragEvent): void => {
    e.preventDefault();
    // 防止子元素触发
    if (e.relatedTarget && !document.contains(e.relatedTarget as Node)) {
      this.dragDropCallback?.onDragLeave?.();
    }
  };

  private handleDrop = async (e: DragEvent): Promise<void> => {
    e.preventDefault();

    const items = e.dataTransfer?.items;
    if (!items) return;

    // 处理文件
    const files: string[] = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.kind === 'file') {
        const file = item.getAsFile();
        if (file) {
          // 创建临时 URL
          files.push(URL.createObjectURL(file));
        }
      }
    }

    if (files.length > 0) {
      this.dragDropCallback?.onFileDrop?.(files);
    }

    // 处理文本
    const text = e.dataTransfer?.getData('text/plain');
    if (text) {
      this.dragDropCallback?.onTextDrop?.(text);
    }
  };

  // 系统主题 - Web 实现
  async getSystemTheme(): Promise<'light' | 'dark' | 'auto'> {
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
      return 'dark';
    }
    return 'light';
  }

  onThemeChange(callback: (theme: 'light' | 'dark') => void): () => void {
    this.themeCallback = callback;

    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');

    const handler = (e: MediaQueryListEvent) => {
      callback(e.matches ? 'dark' : 'light');
    };

    mediaQuery.addEventListener('change', handler);

    return () => {
      mediaQuery.removeEventListener('change', handler);
    };
  }
}

// 导出单例实例
export const webSystemAdapter = new WebSystemAdapter();
