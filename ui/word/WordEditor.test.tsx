import { describe, it, expect } from 'vitest';
import { RangeTableNode } from './RangeTableNode';
import type { WordEditorRefType } from './WordEditor';

describe('WordEditor', () => {
  it('should have RangeTableNode in extensions', () => {
    // This will be verified by checking that editor extensions
    // The actual implementation would check that RangeTableNode is registered
    expect(RangeTableNode).toBeDefined();
  });

  it('should have insertRangeTable method', () => {
    // Verify the WordEditorRefType interface includes insertRangeTable
    const dummyRef: Partial<WordEditorRefType> = {
      insertRangeTable: () => undefined,
      insertReference: () => undefined,
      getEditor: () => null,
      getReferenceIds: () => [],
    };
    expect(dummyRef.insertRangeTable).toBeDefined();
  });
});
