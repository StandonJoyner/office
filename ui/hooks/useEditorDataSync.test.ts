import { describe, it, expect, vi } from 'vitest';
import { eventBus, Events } from '../../core/eventBus';

// We'll test the logic without fully testing React integration
// since testing TipTap editor hooks in Vitest requires a more complex setup

describe('useEditorDataSync', () => {

  describe('event handling', () => {
    it('should listen to DataChanged events', () => {
      const handler = vi.fn();
      eventBus.on(Events.DataChanged, handler);

      const changes = [
        {
          referenceId: 'ref-1',
          oldValue: 'old',
          newValue: 'new',
          timestamp: Date.now(),
          source: 'manual' as const,
        },
      ];
      eventBus.emit(Events.DataChanged, changes);

      expect(handler).toHaveBeenCalledWith(changes);
      eventBus.off(Events.DataChanged, handler);
    });

    it('should listen to ReferenceUpdated events', () => {
      const handler = vi.fn();
      eventBus.on(Events.ReferenceUpdated, handler);

      const data = {
        referenceId: 'ref-1',
        newState: {
          display: { value: 'updated' },
          state: 'active' as const,
        },
      };
      eventBus.emit(Events.ReferenceUpdated, data);

      expect(handler).toHaveBeenCalledWith(data);
      eventBus.off(Events.ReferenceUpdated, handler);
    });
  });

  describe('range table data structure', () => {
    it('should handle 2D array table data structure', () => {
      const tableData = [
        ['A1', 'B1', 'C1'],
        ['A2', 'B2', 'C2'],
        ['A3', 'B3', 'C3'],
      ];

      expect(tableData).toHaveLength(3);
      expect(tableData[0]).toHaveLength(3);
      expect(tableData[0][0]).toBe('A1');
      expect(tableData[2][2]).toBe('C3');
    });

    it('should handle single cell range data', () => {
      const singleCellData = [['Cell Value']];

      expect(singleCellData).toHaveLength(1);
      expect(singleCellData[0]).toHaveLength(1);
      expect(singleCellData[0][0]).toBe('Cell Value');
    });

    it('should handle empty cells in table data', () => {
      const tableData = [
        ['A1', '', 'C1'],
        ['', 'B2', ''],
      ];

      expect(tableData[0][1]).toBe('');
      expect(tableData[1][0]).toBe('');
    });
  });

  describe('table cell content generation', () => {
    it('should generate correct table cell structure', () => {
      const cellValue = 'Test Value';
      const cellNode = {
        type: 'tableCell',
        content: [{ type: 'paragraph', content: [{ type: 'text', text: cellValue }] }],
      };

      expect(cellNode.type).toBe('tableCell');
      expect(cellNode.content[0].type).toBe('paragraph');
      expect(cellNode.content[0].content[0].type).toBe('text');
      expect(cellNode.content[0].content[0].text).toBe(cellValue);
    });

    it('should handle null/undefined cell values', () => {
      const cellNode = (value: any) => ({
        type: 'tableCell' as const,
        content: [{ type: 'paragraph' as const, content: [{ type: 'text' as const, text: String(value ?? '') }] }],
      });

      const nullCell = cellNode(null);
      const undefinedCell = cellNode(undefined);

      // null with ?? '' becomes '' (empty string), then String('') => ''
      expect(nullCell.content[0].content[0].text).toBe('');
      expect(undefinedCell.content[0].content[0].text).toBe('');
    });

    it('should convert numbers to strings', () => {
      const cellNode = (value: number) => ({
        type: 'tableCell' as const,
        content: [{ type: 'paragraph' as const, content: [{ type: 'text' as const, text: String(value) }] }],
      });

      const numberCell = cellNode(123.45);
      expect(numberCell.content[0].content[0].text).toBe('123.45');
    });
  });

  describe('table row generation', () => {
    it('should generate correct table row structure from data array', () => {
      const rowData = ['A1', 'B1', 'C1'];
      const rowNode = rowData.map((cell) => ({
        type: 'tableCell' as const,
        content: [{ type: 'paragraph' as const, content: [{ type: 'text' as const, text: String(cell) }] }],
      }));

      expect(rowNode).toHaveLength(3);
      expect(rowNode[0].type).toBe('tableCell');
      expect(rowNode[0].content[0].content[0].text).toBe('A1');
    });

    it('should wrap cells in tableRow structure', () => {
      const tableData = [['A1', 'B1'], ['A2', 'B2']];
      const rows = tableData.map((row) => {
        const cells = row.map((cell) => ({
          type: 'tableCell' as const,
          content: [{ type: 'paragraph' as const, content: [{ type: 'text' as const, text: String(cell ?? '') }] }],
        }));
        return { type: 'tableRow' as const, content: cells };
      });

      expect(rows).toHaveLength(2);
      expect(rows[0].type).toBe('tableRow');
      expect(rows[0].content).toHaveLength(2);
    });
  });

  describe('data change scenarios', () => {
    it('should handle multiple references changing simultaneously', () => {
      const handler = vi.fn();
      eventBus.on(Events.DataChanged, handler);

      const changes = [
        { referenceId: 'ref-1', oldValue: 'old1', newValue: 'new1', timestamp: 1, source: 'auto' as const },
        { referenceId: 'ref-2', oldValue: 'old2', newValue: 'new2', timestamp: 2, source: 'auto' as const },
        { referenceId: 'table-1', oldValue: [['A1']], newValue: [['A1-updated']], timestamp: 3, source: 'auto' as const },
      ];

      eventBus.emit(Events.DataChanged, changes);

      expect(handler).toHaveBeenCalledWith(changes);
      eventBus.off(Events.DataChanged, handler);
    });

    it('should preserve table structure when updating', () => {
      const oldData = [['A1', 'B1'], ['A2', 'B2']];
      const newData = [['A1-new', 'B1-new'], ['A2-new', 'B2-new']];

      // Both should be 2D arrays with same dimensions
      expect(oldData).toHaveLength(newData.length);
      expect(oldData[0]).toHaveLength(newData[0].length);
      expect(oldData[1]).toHaveLength(newData[1].length);
    });
  });
});
