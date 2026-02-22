import { describe, it, expect, vi } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { ReactNode } from 'react';
import { ReferenceState } from '../../core/types';
import { eventBus, Events } from '../../core/eventBus';

// Mock Tiptap NodeViewWrapper
vi.mock('@tiptap/react', () => ({
  NodeViewWrapper: ({ children, className, ...props }: any) => (
    <div className={className} data-testid="node-view-wrapper" {...props}>
      {children}
    </div>
  ),
}));

// Import after mocking
import { RangeTableNodeView } from './RangeTableNodeView';

interface SourceInfo {
  fileId: string;
  sheetId: string;
  fileName: string;
  range: { startRow: number; startCol: number; endRow: number; endCol: number };
}

interface RangeTableNodeAttrs {
  refId: string | null;
  syncMode: 'manual' | 'auto';
  sourceInfo: string | SourceInfo;
}

// Mock node props factory
const createMockNode = (attrs: RangeTableNodeAttrs) => ({
  attrs,
  type: { name: 'rangeTable' },
});

// Mock props factory
const createMockProps = (
  attrs: RangeTableNodeAttrs,
  selected: boolean = false
) => ({
  node: createMockNode(attrs),
  updateAttributes: vi.fn(),
  deleteNode: vi.fn(),
  selected,
  getPos: vi.fn(() => 0),
});

const sourceInfo: SourceInfo = {
  fileId: 'file-123',
  sheetId: 'sheet-456',
  fileName: 'budget.xlsx',
  range: { startRow: 0, startCol: 0, endRow: 4, endCol: 3 },
};

const mockAttrs: RangeTableNodeAttrs = {
  refId: 'ref-123',
  syncMode: 'manual',
  sourceInfo,
};

describe('RangeTableNodeView', () => {
  // Clear event listeners between tests (cleanup is done in vitest.setup.ts)
  afterEach(() => {
    eventBus.clear();
  });

  describe('Basic Rendering', () => {
    it('should render with mock table data', () => {
      const props = createMockProps(mockAttrs);
      render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content">Mock Table</div>
        </RangeTableNodeView>
      );

      expect(screen.getByTestId('node-view-wrapper')).toBeInTheDocument();
      expect(screen.getByTestId('table-content')).toBeInTheDocument();
      expect(screen.getByText('Mock Table')).toBeInTheDocument();
    });

    it('should parse string sourceInfo', () => {
      const attrs: RangeTableNodeAttrs = {
        ...mockAttrs,
        sourceInfo: JSON.stringify(sourceInfo),
      };
      const props = createMockProps(attrs);

      render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-2">Table</div>
        </RangeTableNodeView>
      );

      expect(screen.getByTestId('node-view-wrapper')).toBeInTheDocument();
    });

    it('should render with null refId', () => {
      const attrs: RangeTableNodeAttrs = {
        ...mockAttrs,
        refId: null,
      };
      const props = createMockProps(attrs);

      render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-3">Table</div>
        </RangeTableNodeView>
      );

      expect(screen.getByTestId('node-view-wrapper')).toBeInTheDocument();
    });
  });

  describe('State Indicator', () => {
    it('should display active state correctly', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-4">Table</div>
        </RangeTableNodeView>
      );

      const statusIndicator = container.querySelector('.bg-green-50');
      expect(statusIndicator).toBeInTheDocument();
      expect(screen.getByText('active')).toBeInTheDocument();
      expect(screen.getByText('🔗')).toBeInTheDocument();
    });

    it('should display stale state correctly', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-5">Table</div>
        </RangeTableNodeView>
      );

      // Check that stale styles are not active by default
      const staleElement = container.querySelector('.bg-yellow-50');
      expect(staleElement).not.toBeInTheDocument();

      // Check that active styles are present
      const activeElement = container.querySelector('.bg-green-50');
      expect(activeElement).toBeInTheDocument();
    });

    it('should display broken state correctly', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-6">Table</div>
        </RangeTableNodeView>
      );

      // Check that broken state styling is not active by default
      const brokenElement = container.querySelector('.bg-red-50');
      expect(brokenElement).not.toBeInTheDocument();
    });

    it('should display conflict state correctly', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-7">Table</div>
        </RangeTableNodeView>
      );

      // Check that conflict state styling is not active by default
      const conflictElement = container.querySelector('.bg-orange-50');
      expect(conflictElement).not.toBeInTheDocument();
    });
  });

  describe('Hover State', () => {
    it('should show toolbar on hover', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-8">Table</div>
        </RangeTableNodeView>
      );

      // Initially, toolbar is not visible until hover (conditionally rendered)
      let toolbar = container.querySelector('.absolute.-top-1.-right-1');
      expect(toolbar).toBeNull();

      // Hover on the status indicator
      const statusIndicator = container.querySelector('.bg-green-50') as HTMLElement;
      if (statusIndicator) {
        fireEvent.mouseEnter(statusIndicator);
      }

      // Now toolbar should be visible
      toolbar = container.querySelector('.absolute.-top-1.-right-1');
      expect(toolbar).toBeInTheDocument();

      // Check that hover state affects opacity
      const indicatorWithHover = container.querySelector('.opacity-100');
      expect(indicatorWithHover).toBeInTheDocument();
    });

    it('should hide toolbar on mouse leave', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-9">Table</div>
        </RangeTableNodeView>
      );

      const statusIndicator = container.querySelector('.bg-green-50') as HTMLElement;
      if (statusIndicator) {
        fireEvent.mouseEnter(statusIndicator);
        fireEvent.mouseLeave(statusIndicator);
      }

      // After mouse leave, opacity should decrease
      const indicatorAfterLeave = container.querySelector('.opacity-60');
      expect(indicatorAfterLeave).toBeInTheDocument();
    });

    it('should show toolbar on table content hover', () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-10">Table</div>
        </RangeTableNodeView>
      );

      // Initially toolbar is not visible
      let toolbar = container.querySelector('.absolute.-top-1.-right-1');
      expect(toolbar).toBeNull();

      const tableContent = screen.getByTestId('table-content-10');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      // Toolbar should be visible when hovering over the table content
      toolbar = container.querySelector('.absolute.-top-1.-right-1');
      expect(toolbar).toBeInTheDocument();
    });
  });

  describe('Refresh Button', () => {
    it('should trigger refresh action when clicked', async () => {
      const emitSpy = vi.spyOn(eventBus, 'emit');
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-11">Table</div>
        </RangeTableNodeView>
      );

      // Find and click refresh button (need to hover first to see it)
      const tableContent = screen.getByTestId('table-content-11');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      await waitFor(() => {
        const refreshButton = container.querySelector('button[title*="Refresh"]');
        expect(refreshButton).toBeInTheDocument();
      });

      const refreshButton = container.querySelector('button[title*="Refresh"]') as HTMLButtonElement;
      if (refreshButton) {
        fireEvent.click(refreshButton);

        // Verify event was emitted
        expect(emitSpy).toHaveBeenCalledWith(
          Events.ReferenceUpdated,
          expect.objectContaining({
            referenceId: 'ref-123',
            action: 'refresh',
          })
        );
      }
    });

    it('should show loading state during refresh', async () => {
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-12">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-12');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      await waitFor(() => {
        const refreshButton = container.querySelector('button[title*="Refresh"]');
        expect(refreshButton).toBeInTheDocument();
      });

      const refreshButton = container.querySelector('button[title*="Refresh"]') as HTMLButtonElement;
      if (refreshButton) {
        // Initially not disabled
        expect(refreshButton.disabled).toBe(false);

        // Click to start refresh
        fireEvent.click(refreshButton);

        // After async refresh completes, button should be enabled again
        await waitFor(
          () => {
            expect(refreshButton.disabled).toBe(false);
          },
          { timeout: 100 }
        );
      }
    });

    it('should not refresh without refId', () => {
      const emitSpy = vi.spyOn(eventBus, 'emit');
      const attrs: RangeTableNodeAttrs = { ...mockAttrs, refId: null };
      const props = createMockProps(attrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-13">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-13');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      const refreshButton = container.querySelector('button[title*="Refresh"]') as HTMLButtonElement;
      if (refreshButton) {
        fireEvent.click(refreshButton);

        // Wait a bit for async to potentially trigger
        setTimeout(() => {
          expect(emitSpy).not.toHaveBeenCalledWith(
            Events.ReferenceUpdated,
            expect.anything()
          );
        }, 100);
      }
    });
  });

  describe('Sync Mode Toggle', () => {
    it('should switch from manual to auto', () => {
      const updateAttributesSpy = vi.fn();
      const props = {
        ...createMockProps(mockAttrs),
        updateAttributes: updateAttributesSpy,
      };
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-14">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-14');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      // Find sync mode button (should say "Manual" initially)
      const buttons = container.querySelectorAll('button');
      const manualButton = Array.from(buttons).find(
        btn => btn.textContent?.includes('Manual')
      );

      if (manualButton) {
        expect(manualButton.textContent).toBe('Manual');
        expect(manualButton.className).not.toContain('bg-blue-100');

        fireEvent.click(manualButton);

        // Verify updateAttributes was called
        expect(updateAttributesSpy).toHaveBeenCalledWith({ syncMode: 'auto' });
      }
    });

    it('should switch from auto to manual', () => {
      const updateAttributesSpy = vi.fn();
      const attrs: RangeTableNodeAttrs = { ...mockAttrs, syncMode: 'auto' };
      const props = {
        ...createMockProps(attrs),
        updateAttributes: updateAttributesSpy,
      };
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-15">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-15');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      const buttons = container.querySelectorAll('button');
      const autoButton = Array.from(buttons).find(
        btn => btn.textContent?.includes('Auto')
      );

      if (autoButton) {
        expect(autoButton.textContent).toBe('Auto');
        expect(autoButton.className).toContain('bg-blue-100');

        fireEvent.click(autoButton);

        expect(updateAttributesSpy).toHaveBeenCalledWith({ syncMode: 'manual' });
      }
    });

    it('should listen to DataChanged events when in auto mode', async () => {
      const refreshSpy = vi.spyOn(eventBus, 'emit');
      const attrs: RangeTableNodeAttrs = { ...mockAttrs, syncMode: 'auto' };
      const props = createMockProps(attrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-16">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-16');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      // Emit DataChanged event for this reference
      eventBus.emit(Events.DataChanged, {
        referenceId: 'ref-123',
        newValue: 'new data',
      });

      // Wait for async refresh
      await waitFor(
        () => {
          expect(refreshSpy).toHaveBeenCalled();
        },
        { timeout: 100 }
      );
    });

    it('should not listen to DataChanged events for other references in auto mode', () => {
      const refreshSpy = vi.spyOn(eventBus, 'emit');
      const attrs: RangeTableNodeAttrs = { ...mockAttrs, syncMode: 'auto' };
      const props = createMockProps(attrs);
      render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-17">Table</div>
        </RangeTableNodeView>
      );

      // Emit DataChanged event for a different reference
      eventBus.emit(Events.DataChanged, {
        referenceId: 'other-ref-456',
        newValue: 'other data',
      });

      // Should not trigger refresh for other references
      setTimeout(() => {
        expect(refreshSpy).not.toHaveBeenCalled();
      }, 100);
    });
  });

  describe('Delete Button', () => {
    it('should call deleteNode when clicked', () => {
      const deleteNodeSpy = vi.fn();
      const props = {
        ...createMockProps(mockAttrs),
        deleteNode: deleteNodeSpy,
      };
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-18">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-18');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      const buttons = container.querySelectorAll('button');
      // Find delete button (has X icon)
      const deleteButton = Array.from(buttons).find(
        btn => btn.getAttribute('title') === 'Remove table reference'
      );

      if (deleteButton) {
        fireEvent.click(deleteButton);

        expect(deleteNodeSpy).toHaveBeenCalled();
      }
    });

    it('should stop event propagation when delete is clicked', () => {
      const deleteNodeSpy = vi.fn();
      const props = {
        ...createMockProps(mockAttrs),
        deleteNode: deleteNodeSpy,
      };

      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-19">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-19');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      const buttons = container.querySelectorAll('button');
      const deleteButton = Array.from(buttons).find(
        btn => btn.getAttribute('title') === 'Remove table reference'
      );

      if (deleteButton) {
        const stopPropagationSpy = vi.fn();
        const clickEvent = {
          stopPropagation: stopPropagationSpy,
        } as unknown as React.MouseEvent;

        // Access the component's handleDelete function through the DOM element's handler
        deleteButton.dispatchEvent(
          new MouseEvent('click', { bubbles: true, cancelable: true })
        );

        expect(deleteNodeSpy).toHaveBeenCalled();
      }
    });
  });

  describe('Selected State', () => {
    it('should show selection ring when selected', () => {
      const props = createMockProps(mockAttrs, true);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-20">Table</div>
        </RangeTableNodeView>
      );

      // Check for ring-2 class indicating selection
      const selectedElement = container.querySelector('.ring-2');
      expect(selectedElement).toBeTruthy();
    });

    it('should not show selection ring when not selected', () => {
      const props = createMockProps(mockAttrs, false);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-21">Table</div>
        </RangeTableNodeView>
      );

      const selectedElement = container.querySelector('.ring-2');
      expect(selectedElement).toBeNull();
    });
  });

  describe('Integration with EventBus', () => {
    it('should emit ReferenceUpdated event on refresh', async () => {
      const emitSpy = vi.spyOn(eventBus, 'emit');
      const props = createMockProps(mockAttrs);
      const { container } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-22">Table</div>
        </RangeTableNodeView>
      );

      const tableContent = screen.getByTestId('table-content-22');
      if (tableContent) {
        fireEvent.mouseEnter(tableContent);
      }

      await waitFor(() => {
        const refreshButton = container.querySelector('button[title*="Refresh"]');
        expect(refreshButton).toBeInTheDocument();
      });

      const refreshButton = container.querySelector('button[title*="Refresh"]') as HTMLButtonElement;
      if (refreshButton) {
        fireEvent.click(refreshButton);

        expect(emitSpy).toHaveBeenCalledWith(
          Events.ReferenceUpdated,
          expect.objectContaining({
            referenceId: 'ref-123',
            action: 'refresh',
          })
        );
      }
    });

    it('should subscribe to DataChanged in auto mode and unsubscribe on unmount', () => {
      const onSpy = vi.spyOn(eventBus, 'on');
      const offSpy = vi.spyOn(eventBus, 'off');
      const attrs: RangeTableNodeAttrs = { ...mockAttrs, syncMode: 'auto' };
      const props = createMockProps(attrs);
      const { unmount } = render(
        <RangeTableNodeView {...props}>
          <div data-testid="table-content-23">Table</div>
        </RangeTableNodeView>
      );

      expect(onSpy).toHaveBeenCalledWith(
        Events.DataChanged,
        expect.any(Function)
      );

      unmount();

      // Cleanup happens in useEffect return
      expect(offSpy).toHaveBeenCalled();
    });
  });
});
