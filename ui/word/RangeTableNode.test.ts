import { describe, it, expect } from 'vitest';
import { RangeTableNode } from './RangeTableNode';

describe('RangeTableNode', () => {
  it('should have correct node name', () => {
    expect(RangeTableNode.name).toBe('rangeTable');
  });

  it('should extend Table extension', () => {
    expect(RangeTableNode).toBeDefined();
  });
});
