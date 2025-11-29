# @ricsam/formula-engine

## 0.0.17

### Patch Changes

- - Add cut-paste reference updates, multi-area style support, and proper overlapping range handling

    **Breaking Changes:**

    - `ConditionalStyle` and `DirectCellStyle` now use `areas: RangeAddress[]` instead of `area: RangeAddress`
      - Enables Excel-like multi-area styles when cutting cells from styled regions
      - Cutting a cell from a styled area now punches a hole, creating multiple areas in the same style object

    **New Features:**

    - Cell and range references are automatically updated when cells are cut and pasted
      - When a cell (e.g., A1) is cut to a new location (e.g., D5), all formulas referencing A1 update to D5
      - Both relative and absolute references are updated (e.g., `=$A$1` becomes `=$D$5`)
      - Range references update only when the entire range is moved (e.g., `=SUM(A1:B2)` → `=SUM(D5:E6)`)
      - Partial range moves don't update references (prevents breaking formulas)
      - Works across sheets and workbooks
    - Added `moveCell(source, target)` method - programmatically move cells with reference updates
    - Added `moveRange(sourceRange, target)` method - programmatically move ranges with reference updates
    - Added `getAllCellStyles()` and `getAllConditionalStyles()` methods to engine for testing/serialization
    - Cut operations always use `pasteCells()` not `fillAreas()` (Excel behavior)

    **Style System Improvements:**

    - **Multi-area style support**: Styles can now cover non-contiguous ranges
      - Cutting cell L11 from styled area J3:N20 creates 4 areas: J3:N10, J12:N20, J11:K11, M11:N11
      - Hole-punching keeps all areas in the same style object (not separate styles)
    - `clearCellStylesInRange()` now properly punches holes instead of creating new style objects
    - All style operations (filtering, updating, rendering) now work with multiple areas

    **Architecture Improvements:**

    - Restructured cut flow: snapshot → remove → update references → apply (proper Excel-like behavior)
    - Renamed internal `copyCells()` to `pasteCells()` in CopyManager for consistency
    - Enhanced `CellSnapshot` to include style information for proper cut operations
    - Added `WorkbookManager.updateFormulasExcluding()` to update references excluding moved cells
    - All sheet modifications now go through proper WorkbookManager APIs (maintains indexes)
    - Snapshot-based copying handles overlapping source/target ranges correctly
    - Formulas in moved cells are adjusted once; formulas referencing moved cells updated separately

    **Bug Fixes:**

    - Fixed overlapping cut operations (e.g., cutting A1:B4 to B1:C4 now correctly clears A1-A4)
    - Fixed cache invalidation when modifying cells during cut operations
    - Cut cells now properly carry their styles to the new location

    **Implementation Details:**

    - New `src/core/cell-mover.ts` module with formula transformation utilities
    - Split CopyManager into `cutCells()` (move flow) and `copyOnlyCells()` (copy flow)
    - Cut and move operations use identical underlying logic
    - Comprehensive test coverage: 563 tests passing, including overlapping ranges and multi-area styles

## 0.0.16

### Patch Changes

- fix clone workbook

## 0.0.15

### Patch Changes

- improve color management

## 0.0.14

### Patch Changes

- 7c967ac: add smartPaste

## 0.0.13

### Patch Changes

- enable copy of value only or style only

## 0.0.12

### Patch Changes

- add getStyleForRange

## 0.0.11

### Patch Changes

- make getCellStyles take a range

## 0.0.10

### Patch Changes

- Update range/copy-manager logic and tests

## 0.0.9

### Patch Changes

- add engine.clearCellStyles()

## 0.0.8

### Patch Changes

- add copyCells method and expand styling capabilities

## 0.0.7

### Patch Changes

- export color utils

## 0.0.6

### Patch Changes

- add conditional formatting and cell styling support

## 0.0.5

### Patch Changes

- remove last index.ts file

## 0.0.4

### Patch Changes

- fix base import path

## 0.0.3

### Patch Changes

- export utils

## 0.0.2

### Patch Changes

- add lib entry

## 0.0.1

### Patch Changes

- initial release
