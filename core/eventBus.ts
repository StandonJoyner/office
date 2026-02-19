export enum Events {
  ReferenceCreated = 'reference:created',
  ReferenceUpdated = 'reference:updated',
  ReferenceDeleted = 'reference:deleted',
  ReferenceBroken = 'reference:broken',
  DataChanged = 'data:changed',
  DocumentSaved = 'document:saved',
  ExcelWorkbookLoaded = 'excel:workbook:loaded',
  ExcelCellSelected = 'excel:cell:selected',
  ExcelRangeSelected = 'excel:range:selected',
}

export type EventHandler<T = any> = (data: T) => void;

type EventMap = {
  [K in Events]: EventHandler<any>;
};

class EventBus {
  private listeners: Map<Events, Set<EventHandler>> = new Map();

  on<T = any>(event: Events, handler: EventHandler<T>): void {
    if (!this.listeners.has(event)) {
      this.listeners.set(event, new Set());
    }
    this.listeners.get(event)!.add(handler);
  }

  off<T = any>(event: Events, handler: EventHandler<T>): void {
    const handlers = this.listeners.get(event);
    if (handlers) {
      handlers.delete(handler);
      if (handlers.size === 0) {
        this.listeners.delete(event);
      }
    }
  }

  emit<T = any>(event: Events, payload: T): void {
    const handlers = this.listeners.get(event);
    if (handlers) {
      handlers.forEach((handler) => {
        try {
          handler(payload);
        } catch (error) {
          console.error(`Error in event handler for ${event}:`, error);
        }
      });
    }
  }

  clear(): void {
    this.listeners.clear();
  }
}

export const eventBus = new EventBus();
