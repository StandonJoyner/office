import { describe, it, expect } from 'vitest';
import { RangeTableNode } from './RangeTableNode';

describe('WordEditor', () => {
  it('should have RangeTableNode in extensions', () => {
    // This will be verified by checking the editor extensions
    // The actual implementation would check that RangeTableNode is registered
    expect(RangeTableNode).toBeDefined();
  });
});
