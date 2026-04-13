# @ricsam/formula-engine

## 0.2.7

### Patch Changes

- better cache invalidation

## 0.2.6

### Patch Changes

- bump snapshot version

## 0.2.5

### Patch Changes

- Fixed a dependency-resolution bug where open ranges could keep stable frontier candidates alive for too long during reevaluation.

  In the failing cases, a formula would reference an open range such as `Q18:Q` or `A1:INFINITY`. The range planner correctly added frontier candidates that might spill into that range, but once those candidates had been evaluated and were known not to spill, they were not being discarded early enough. That left avoidable transient dependencies in the graph, caused repeated evaluation-plan rebuilds, and made some operations such as toggling a table feel much slower than they should.

  This change lets range nodes prune frontier candidates as soon as those candidates are stable enough to prove that they do not spill into the range, and it excludes discarded frontier dependencies from traversal of the active dependency graph. In practice this removes a large amount of false-positive reevaluation work for open ranges and significantly improves recalculation time after table and range-shape changes.

  While fixing that, the reevaluation path for circular open-range cases was also tightened so a formula such as `SUM(A1:INFINITY)` does not count its own circular result back into the aggregate on a later rerun.

## 0.2.4

### Patch Changes

- sheet rename fixes

## 0.2.3

### Patch Changes

- snapshot whole engine + performant edits

## 0.2.2

### Patch Changes

- add addGridSchema method to the engine class

## 0.2.1

### Patch Changes

- fix type issue

## 0.2.0

### Minor Changes

- ### Added

  - **Grid Schema Support**: Added `addGridSchema()` method to define schemas for 2D ranges of cells. Grid schemas provide readonly `columns` and `rows` getters for column-major and row-major array access, plus `setValue()` and `getValue()` methods for individual cell access.

  ### Changed

  - **Explicit Write Functions**: Schema definitions now use explicit `write` functions instead of guessing how to serialize parsed values. This provides better type safety and control over serialization.
  - **Optional Write for Primitive Types**: The `write` function is now optional when `parse` returns a `SerializedCellValue` (number, string, or boolean). For complex types (objects), `write` is required and TypeScript will error if omitted.
  - **GridOrm API**: Removed `columns` and `rows` setters from GridOrm. Use `setValue(value, position)` instead. The getters now return readonly arrays.
  - **Table Headers**: Updated `defineHeader()` helper to support optional `write` functions. Headers now use explicit write functions for serialization.

  ### Migration Guide

  **Grid Schemas:**

  ```typescript
  // Before: No grid schema support
  // After:
  .addGridSchema(
    "matrix",
    { workbookName: "wb", sheetName: "s" },
    { start: { col: 0, row: 0 }, end: { col: 9, row: 9 } },
    (value) => parseNumber(value)
    // write optional for primitive types
  )

  // Access grid data
  schema.matrix.columns  // readonly number[][]
  schema.matrix.rows     // readonly number[][]
  schema.matrix.setValue(42, { col: 0, row: 0 })
  schema.matrix.getValue({ col: 0, row: 0 })
  ```

  **Cell Schemas with Complex Types:**

  ```typescript
  // Before: Guessing serialization
  .addCellSchema("item", addr, (v) => ({ value: v, meta: true }))

  // After: Explicit write function required
  .addCellSchema(
    "item",
    addr,
    (v) => ({ value: v, meta: true }),
    (item) => ({ value: item.value })  // REQUIRED for object types
  )
  ```

  **Table Headers:**

  ```typescript
  // Before: Simple header definition
  { value: { parse: parseNumber, index: 0 } }

  // After: Use defineHeader helper
  { value: defineHeader(0, parseNumber) }
  // Or with custom write:
  { value: defineHeader(0, parseNumber, (v) => ({ value: v })) }
  ```

### Patch Changes

- export defineHeader

## 0.1.0

### Minor Changes

- ### Added

  - **Grid Schema Support**: Added `addGridSchema()` method to define schemas for 2D ranges of cells. Grid schemas provide readonly `columns` and `rows` getters for column-major and row-major array access, plus `setValue()` and `getValue()` methods for individual cell access.

  ### Changed

  - **Explicit Write Functions**: Schema definitions now use explicit `write` functions instead of guessing how to serialize parsed values. This provides better type safety and control over serialization.
  - **Optional Write for Primitive Types**: The `write` function is now optional when `parse` returns a `SerializedCellValue` (number, string, or boolean). For complex types (objects), `write` is required and TypeScript will error if omitted.
  - **GridOrm API**: Removed `columns` and `rows` setters from GridOrm. Use `setValue(value, position)` instead. The getters now return readonly arrays.
  - **Table Headers**: Updated `createHeader()` helper to support optional `write` functions. Headers now use explicit write functions for serialization.

  ### Migration Guide

  **Grid Schemas:**

  ```typescript
  // Before: No grid schema support
  // After:
  .addGridSchema(
    "matrix",
    { workbookName: "wb", sheetName: "s" },
    { start: { col: 0, row: 0 }, end: { col: 9, row: 9 } },
    (value) => parseNumber(value)
    // write optional for primitive types
  )

  // Access grid data
  schema.matrix.columns  // readonly number[][]
  schema.matrix.rows     // readonly number[][]
  schema.matrix.setValue(42, { col: 0, row: 0 })
  schema.matrix.getValue({ col: 0, row: 0 })
  ```

  **Cell Schemas with Complex Types:**

  ```typescript
  // Before: Guessing serialization
  .addCellSchema("item", addr, (v) => ({ value: v, meta: true }))

  // After: Explicit write function required
  .addCellSchema(
    "item",
    addr,
    (v) => ({ value: v, meta: true }),
    (item) => ({ value: item.value })  // REQUIRED for object types
  )
  ```

  **Table Headers:**

  ```typescript
  // Before: Simple header definition
  { value: { parse: parseNumber, index: 0 } }

  // After: Use createHeader helper
  { value: createHeader(0, parseNumber) }
  // Or with custom write:
  { value: createHeader(0, parseNumber, (v) => ({ value: v })) }
  ```

## 0.0.22

### Patch Changes

- Rename API to Schema throughout the codebase. All API-related types, functions, and properties have been renamed to use "schema" terminology. `defineApi` is now `defineSchema`, `engine.api` is now `engine.schema`, and all related types have been updated. The `createApi` alias has been removed. Added runtime schema methods `addTableSchema` and `addCellSchema` to FormulaEngine for dynamically registering schemas after engine creation. Schema validation now works correctly for both initial and runtime-added schemas.

## 0.0.21

### Patch Changes

- Remove custom methods layer from API. The `methods` parameter has been removed from `addTableApi` and `addCellApi`. ORM methods (`findWhere`, `findAllWhere`, `append`, `updateWhere`, `removeWhere`, `count` for tables and `read`, `write`, `getAddress` for cells) are now exposed directly on the API namespace.

## 0.0.20

### Patch Changes

- Implement Excel-like table spill restrictions: spilling formulas inside tables and formulas that would spill into table areas now return #SPILL! error instead of succeeding.

## 0.0.19

### Patch Changes

- fix exported lib

## 0.0.18

### Patch Changes

- api support and new state management

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
