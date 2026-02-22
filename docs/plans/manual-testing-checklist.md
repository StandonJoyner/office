# Manual Testing Checklist - Range Reference Feature

This document provides a comprehensive checklist for manual verification testing of the Excel range reference table functionality in Word documents.

## Prerequisites

- [ ] Development server running (`npm run dev`)
- [ ] Both Word and Excel editors loaded
- [ ] Sample Excel data available with a range of cells

---

## 1. Table Insertion

### 1.1 Toolbar Visibility
- [ ] **Test**: Locate the Word editor toolbar
- [ ] **Expected**: A "+ Table Reference" button is visible in the toolbar

### 1.2 Button Click Trigger
- [ ] **Test**: Click the "+ Table Reference" button
- [ ] **Expected**: The button triggers the `onInsertTableReference` callback

### 1.3 Table Insertion with Valid Data
- [ ] **Test**: Simulate callback returning valid range data:
  ```typescript
  {
    source: {
      workbookId: 'test-workbook',
      sheetId: 'sheet1',
      range: { startRow: 0, startCol: 0, endRow: 2, endCol: 2 }
    },
    data: [
      ['A1', 'B1', 'C1'],
      ['A2', 'B2', 'C2'],
      ['A3', 'B3', 'C3']
    ],
    metadata: { updated: Date.now() }
  }
  ```
- [ ] **Expected**: A table is inserted at cursor position with:
  - Correct data in all cells
  - Proper header row styling (gray background)
  - Borders on all cells
  - Reference node wrapper with state indicator

### 1.4 Table Insertion with Invalid Data
- [ ] **Test**: Simulate callback returning null/undefined
- [ ] **Expected**: No table is inserted, editor continues normally

---

## 2. Table Display

### 2.1 Basic Table Structure
- [ ] **Test**: Inspect inserted table
- [ ] **Expected**: Table displays with:
  - Correct number of rows (data rows + 1 header row)
  - Correct number of columns
  - Data properly aligned in cells
  - Header row with gray background (`bg-gray-100`)
  - All cells with borders (`border` class)

### 2.2 State Indicator - Active State
- [ ] **Test**: Insert table with fresh data (timestamp recent)
- [ ] **Expected**: State indicator shows green dot with "Active" label

### 2.3 State Indicator - Stale State
- [ ] **Test**: Modify table's `metadata.updated` to an old timestamp
- [ ] **Expected**: State indicator shows yellow dot with "Stale" label

### 2.4 State Indicator - Broken State
- [ ] **Test**: Set table's state to 'broken'
- [ ] **Expected**: State indicator shows red dot with "Broken" label

### 2.5 State Indicator - Conflict State
- [ ] **Test**: Set table's state to 'conflict'
- [ ] **Expected**: State indicator shows orange dot with "Conflict" label

### 2.6 Source Information Storage
- [ ] **Test**: Check table's stored data structure
- [ ] **Expected**: Table node contains:
  - `source.workbookId` set correctly
  - `source.sheetId` set correctly
  - `source.range` with correct coordinates
  - `data` array with cell values
  - `metadata.updated` timestamp

### 2.7 Source Information Display
- [ ] **Test**: Inspect the source info panel in ReferenceNodeView
- [ ] **Expected**: Displays:
  - Workbook ID
  - Sheet ID
  - Range coordinates (e.g., "A1:C3")
  - Last updated timestamp

---

## 3. Sync Modes

### 3.1 Initial Sync Mode
- [ ] **Test**: Check initial sync mode state
- [ ] **Expected**: Defaults to 'manual' mode

### 3.2 Sync Mode Toggle - Manual to Auto
- [ ] **Test**: Click sync mode toggle button
- [ ] **Expected**: Mode changes to 'auto', button icon changes to "Auto" label

### 3.3 Sync Mode Toggle - Auto to Manual
- [ ] **Test**: Click sync mode toggle button again
- [ ] **Expected**: Mode changes to 'manual', button icon shows manual refresh icon

### 3.4 Manual Sync - Single Table
- [ ] **Test**: Click refresh button on a single table in manual mode
- [ ] **Expected**:
  - Callback is triggered with source information
  - If data changes, table updates with new values
  - State indicator updates appropriately
  - Last updated timestamp refreshes

### 3.5 Auto Sync - Data Changes Trigger Refresh
- [ ] **Test**:
  1. Switch to 'auto' mode
  2. Simulate Excel data change for the referenced range
  3. Wait for auto-sync interval (or trigger manually)
- [ ] **Expected**: Table automatically refreshes with new data

### 3.6 Auto Sync - No Changes
- [ ] **Test**:
  1. Switch to 'auto' mode
  2. Trigger sync when data hasn't changed
- [ ] **Expected**: No visible change in table, state remains 'active'

### 3.7 Sync Mode Persistence
- [ ] **Test**: Change sync mode, then reload page
- [ ] **Expected**: Sync mode persists (if using storage)

---

## 4. Table Editing

### 4.1 Hover Shows Toolbar
- [ ] **Test**: Move mouse over a table reference
- [ ] **Expected**: Table toolbar appears above/below the table

### 4.2 Toolbar Visibility - Not Hovering
- [ ] **Test**: Move mouse away from the table
- [ ] **Expected**: Toolbar fades out/disappears

### 4.3 Refresh Button Functionality
- [ ] **Test**: Click refresh button in table toolbar
- [ ] **Expected**:
  - `onRefresh` callback is triggered
  - Table updates with latest data
  - Loading state is shown (if applicable)
  - State indicator updates

### 4.4 Delete Button - Confirmation
- [ ] **Test**: Click delete button in table toolbar
- [ ] **Expected**: Confirmation dialog appears asking to confirm deletion

### 4.5 Delete Button - Confirm
- [ ] **Test**: Confirm deletion in dialog
- [ ] **Expected**:
  - Table is removed from document
  - Cursor is positioned where table was
  - Reference is removed from data link manager

### 4.6 Delete Button - Cancel
- [ ] **Test**: Cancel deletion in dialog
- [ ] **Expected**: Table remains in document, nothing changes

### 4.7 Selection Ring - Select Table
- [ ] **Test**: Click on a table reference
- [ ] **Expected**: Selection ring appears around the table

### 4.8 Selection Ring - Deselect
- [ ] **Test**: Click outside the table
- [ ] **Expected**: Selection ring disappears

### 4.9 Selection Ring - Visual Style
- [ ] **Test**: Inspect selection ring
- [ ] **Expected**: Ring has:
  - Blue border color (`ring-2`, `ring-blue-500`)
  - Slight rounded corners (`rounded-sm`)
  - Appropriate padding/spacing

---

## 5. Data Updates

### 5.1 Excel Data Changes - Auto Mode
- [ ] **Test**:
  1. Insert table reference in auto sync mode
  2. Modify Excel data in the referenced range
  3. Wait for auto-sync trigger
- [ ] **Expected**:
  - Table automatically updates with new data
  - State indicator shows 'active'
  - Last updated timestamp refreshes

### 5.2 Excel Data Changes - Manual Mode
- [ ] **Test**:
  1. Insert table reference in manual sync mode
  2. Modify Excel data in the referenced range
- [ ] **Expected**:
  - Table does NOT update automatically
  - State indicator changes to 'stale'
  - User must manually refresh to see changes

### 5.3 "Refresh All" Button - Multiple Tables
- [ ] **Test**:
  1. Insert multiple table references with stale data
  2. Click "Refresh All References" button
- [ ] **Expected**:
  - All tables refresh with latest data
  - All state indicators update to 'active'
  - Loading indicators shown during refresh

### 5.4 "Refresh All" Button - Partial Updates
- [ ] **Test**:
  1. Insert multiple tables, some active and some stale
  2. Click "Refresh All References"
- [ ] **Expected**: All tables attempt refresh, regardless of current state

### 5.5 State Transition - Active to Stale
- [ ] **Test**:
  1. Insert fresh table (active state)
  2. Modify Excel data without refreshing
- [ ] **Expected**: State changes from 'active' to 'stale'

### 5.6 State Transition - Stale to Active
- [ ] **Test**:
  1. Have a stale table
  2. Refresh table
- [ ] **Expected**: State changes from 'stale' to 'active'

### 5.7 State Transition - Broken
- [ ] **Test**:
  1. Insert table reference
  2. Delete source Excel sheet/workbook
  3. Attempt refresh
- [ ] **Expected**: State changes to 'broken', error message shown

### 5.8 Data Format Preservation
- [ ] **Test**:
  1. Insert table with numeric data
  2. Insert table with text data
  3. Insert table with mixed data
- [ ] **Expected**: Data formats are preserved correctly (numbers, text, dates)

### 5.9 Large Range Handling
- [ ] **Test**: Insert reference to large range (e.g., 20x20 cells)
- [ ] **Expected**:
  - Table displays without layout issues
  - Performance remains acceptable
  - All cells show correct data

---

## 6. Edge Cases

### 6.1 Empty Range
- [ ] **Test**: Insert reference to empty range (0x0 cells)
- [ ] **Expected**: Graceful handling, no errors, may show empty table or warning

### 6.2 Single Cell Range
- [ ] **Test**: Insert reference to single cell (1x1)
- [ ] **Expected**: Table displays with single cell and header

### 6.3 Very Long Text in Cells
- [ ] **Test**: Insert reference with very long text content
- [ ] **Expected**:
  - Text wraps properly
  - Table maintains reasonable width
  - No horizontal overflow issues

### 6.4 Special Characters in Data
- [ ] **Test**: Insert reference with special characters (<, >, &, quotes, etc.)
- [ ] **Expected**: Characters display correctly, no HTML injection

### 6.5 Rapid Refresh Attempts
- [ ] **Test**: Click refresh button multiple times rapidly
- [ ] **Expected**:
  - Requests are debounced or queued
  - No race conditions or duplicate updates
  - Final state is correct

### 6.6 Multiple Tables with Same Source
- [ ] **Test**: Insert multiple references to same Excel range
- [ ] **Expected**:
  - All tables display independently
  - All update when source changes (in auto mode)
  - Each maintains its own state

### 6.7 Concurrent Edits
- [ ] **Test**:
  1. User manually editing a table cell
  2. Auto-sync triggers update
- [ ] **Expected**:
  - Conflict detected
  - State changes to 'conflict'
  - User prompted to resolve

---

## 7. Accessibility

### 7.1 Keyboard Navigation
- [ ] **Test**: Tab through document
- [ ] **Expected**: Tables are focusable with keyboard

### 7.2 Screen Reader Support
- [ ] **Test**: Use screen reader to inspect table
- [ ] **Expected**:
  - Table structure announced correctly
  - State indicator announced
  - Source information available

### 7.3 High Contrast Mode
- [ ] **Test**: Enable high contrast mode
- [ ] **Expected**: Table remains readable with appropriate contrast

### 7.4 Focus Indicators
- [ ] **Test**: Navigate to table with keyboard
- [ ] **Expected**: Clear focus indicator on table/selection ring

---

## 8. Browser Compatibility

### 8.1 Chrome
- [ ] **Test**: Test all functionality in Chrome
- [ ] **Expected**: All features work correctly

### 8.2 Firefox
- [ ] **Test**: Test all functionality in Firefox
- [ ] **Expected**: All features work correctly

### 8.3 Safari
- [ ] **Test**: Test all functionality in Safari
- [ ] **Expected**: All features work correctly

### 8.4 Edge
- [ ] **Test**: Test all functionality in Edge
- [ ] **Expected**: All features work correctly

---

## 9. Performance

### 9.1 Initial Load with Tables
- [ ] **Test**: Load document with 10+ table references
- [ ] **Expected**: Document loads within 2 seconds

### 9.2 Refresh Performance
- [ ] **Test**: Refresh a table with 10x10 range
- [ ] **Expected**: Refresh completes within 500ms

### 9.3 "Refresh All" Performance
- [ ] **Test**: Refresh 10 tables simultaneously
- [ ] **Expected**: All refreshes complete within 2 seconds

### 9.4 Memory Usage
- [ ] **Test**: Monitor memory while working with tables
- [ ] **Expected**: No significant memory leaks after extended use

---

## 10. User Experience

### 10.1 Intuitive Toolbar Placement
- [ ] **Test**: Ask first-time user to find "+ Table Reference" button
- [ ] **Expected**: User can locate button within 10 seconds

### 10.2 Clear State Indicators
- [ ] **Test**: Show table in each state to new user
- [ ] **Expected**: User understands what each state means

### 10.3 Helpful Error Messages
- [ ] **Test**: Trigger error (broken reference)
- [ ] **Expected**: Clear, actionable error message displayed

### 10.4 Smooth Animations
- [ ] **Test**: Trigger various state changes
- [ ] **Expected**: Transitions are smooth and not distracting

---

## Sign-Off

**Tester Name**: ______________________________

**Date**: ______________________________________

**Browser(s) Tested**: ____________________________

**Overall Status**: [ ] Passed [ ] Passed with Notes [ ] Failed

**Notes**:
__________________________________________________________________________
__________________________________________________________________________
__________________________________________________________________________
