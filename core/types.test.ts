import { describe, it, expect } from 'vitest';
import type { RangeTableMeta } from './types';

describe('RangeTableMeta', () => {
  it('should accept valid table metadata', () => {
    const meta: RangeTableMeta = {
      rowCount: 3,
      colCount: 2,
      syncMode: 'manual',
      preserveFormatting: true,
    };
    expect(meta.rowCount).toBe(3);
    expect(meta.colCount).toBe(2);
  });
});
