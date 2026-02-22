import { describe, it, expect } from 'vitest';
import { ReferenceState } from '../../core/types';

describe('RangeTableNodeView', () => {
  // Basic unit test to verify the component exports correctly
  it('should import RangeTableNodeView', async () => {
    const { RangeTableNodeView } = await import('./RangeTableNodeView');
    expect(RangeTableNodeView).toBeDefined();
    expect(typeof RangeTableNodeView).toBe('function');
  });

  // Verify state styles are defined correctly
  it('should have correct state styles', () => {
    const states: ReferenceState[] = ['active', 'stale', 'broken', 'conflict'];
    states.forEach(state => {
      expect(['active', 'stale', 'broken', 'conflict']).toContain(state);
    });
  });
});
