import { describe, it, expect, vi } from 'vitest';
import { eventBus, Events } from './eventBus';

describe('EventBus', () => {
  it('should register and emit events', () => {
    const handler = vi.fn();
    eventBus.on(Events.ReferenceCreated, handler);

    const payload = { id: 'test-ref' };
    eventBus.emit(Events.ReferenceCreated, payload);

    expect(handler).toHaveBeenCalledWith(payload);
  });

  it('should unregister event handlers', () => {
    const handler = vi.fn();
    eventBus.on(Events.ReferenceUpdated, handler);
    eventBus.off(Events.ReferenceUpdated, handler);

    eventBus.emit(Events.ReferenceUpdated, { id: 'test-ref' });

    expect(handler).not.toHaveBeenCalled();
  });

  it('should call all registered handlers for an event', () => {
    const handler1 = vi.fn();
    const handler2 = vi.fn();
    eventBus.on(Events.DataChanged, handler1);
    eventBus.on(Events.DataChanged, handler2);

    eventBus.emit(Events.DataChanged, { referenceId: 'test' });

    expect(handler1).toHaveBeenCalled();
    expect(handler2).toHaveBeenCalled();
  });
});
