# FormulaEngine

A TypeScript-based spreadsheet formula evaluation library designed for high-performance calculation of formulas across sparse datasets.

## Features

- **Sparse-aware architecture** - Only populated cells consume memory
- **Map-based storage** - O(1) cell access with automatic sparse handling
- **Excel-compatible** addressing (A1 notation)
- **Multi-sheet support** with sheet management
- **Named expressions** with global and sheet-level scoping
- **Copy/paste operations** with clipboard support
- **TypeScript-first** design with comprehensive type safety
- **Editor-neutral formula analysis** with semantic tokens, diagnostics, and
  cell/range targets for syntax highlighting and spreadsheet overlays

## Installation

```bash
bun install
```

## Quick Start

```typescript
import { FormulaEngine } from "formula-engine";

// Create a new engine
const engine = FormulaEngine.buildEmpty();

// Add a sheet
const sheetName = engine.addSheet("Sheet1");
const sheetId = engine.getSheetId(sheetName);

// Set cell values
engine.setCellContent({ sheet: sheetId, col: 0, row: 0 }, 42);
engine.setCellContent({ sheet: sheetId, col: 1, row: 0 }, 58);

// Set a formula (evaluation not yet implemented)
engine.setCellContent({ sheet: sheetId, col: 2, row: 0 }, "=A1+B1");

// Get cell value
const value = engine.getCellValue({ sheet: sheetId, col: 0, row: 0 }); // 42

// Set multiple values at once
engine.setCellContent({ sheet: sheetId, col: 0, row: 2 }, [
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9],
]);

// Get range values
const range = {
  start: { sheet: sheetId, col: 0, row: 2 },
  end: { sheet: sheetId, col: 2, row: 4 },
};
const values = engine.getRangeValues(range); // [[1,2,3],[4,5,6],[7,8,9]]
```

## Clone Workbooks And Sheets

```typescript
engine.cloneWorkbook("Workbook1", "Workbook1 Copy");

const clonedSheet = engine.cloneSheet({
  workbookName: "Workbook1",
  sheetName: "Sheet1",
  newSheetName: "Sheet1 Copy",
});
```

`cloneSheet` copies cell content and metadata, sheet metadata, sheet-scoped
named expressions, tables, styles, cell data types, and range metadata. Explicit
self-references are rewritten to the new sheet. Because table names are unique
within a workbook, cloned tables use the first available `_2`, `_3`, and so on
suffix. The clone is added at the end of the workbook's sheet order, and the
entire operation is one undo/redo step.

## Search And Replace Raw Strings

```typescript
engine.addWorkbook("Workbook1");
engine.addSheet({ workbookName: "Workbook1", sheetName: "Sheet1" });
engine.setSheetContent(
  { workbookName: "Workbook1", sheetName: "Sheet1" },
  new Map([
    ["A1", "=SUM(B1:B10)"],
    ["A2", "draft summary"],
  ])
);

const matches = engine.search("sum", { workbookName: "Workbook1" });
const firstTenMatches = engine.search("sum", {
  workbookName: "Workbook1",
  maxResults: 10,
});
const oneChange = engine.replace("sum", "avg", {
  workbookName: "Workbook1",
  sheetName: "Sheet1",
  cellReference: "A1",
  occurrenceIndex: 0,
});
const allChanges = engine.replaceAll("draft", "published", {
  workbookName: "Workbook1",
});

// Works on any stored string cell, including formulas and plain text
// search() returns at most 1,000 matches by default to keep large UI searches
// responsive. Pass maxResults to change the cap, or Number.POSITIVE_INFINITY
// when a batch workflow really needs an unbounded result set.
// [
//   {
//     workbookName: "Workbook1",
//     sheetName: "Sheet1",
//     cellReference: "A1",
//     cellContent: "=SUM(B1:B10)",
//     contentKind: "formula",
//     occurrenceIndex: 0,
//     startIndex: 1,
//     endIndexExclusive: 4,
//     matchedText: "SUM"
//   }
// ]
```

## Undo And Redo

Undo/redo history stores incremental reversible changes rather than copies of
the complete engine. Retention is bounded by both entry count and estimated
memory usage:

```typescript
const engine = FormulaEngine.buildEmpty({
  undoRedo: {
    maxEntries: 100,
    maxBytes: 64 * 1024 * 1024,
  },
});

engine.undo();
engine.redo();

const history = engine.getUndoRedoState();
console.log(history.undoDepth, history.undoBytes);
console.log(history.redoDepth, history.redoBytes);

engine.transact(() => {
  // Synchronous only: grouped into one atomic history entry and one update.
  engine.setCellContent(
    { workbookName: "Book", sheetName: "Sheet1", rowIndex: 0, colIndex: 0 },
    1
  );
  engine.setCellContent(
    { workbookName: "Book", sheetName: "Sheet1", rowIndex: 0, colIndex: 1 },
    2
  );
});
```

The defaults are 100 entries and 64 MiB. If one mutation is larger than the
configured byte budget, the engine clears the existing history and does not
retain that mutation. This history barrier prevents a later undo from crossing
an operation that could not be recorded safely.

History detaches and accounts for primitives, plain objects, arrays, ordered
`Map` values with primitive keys, `Set` values with primitive elements, dates,
regular expressions, errors, and binary buffers. Object identity inside
metadata is intentionally not preserved across undo/redo.
Metadata containing functions, accessors, weak collections, promises, custom
class instances, arbitrary-precision `bigint` values, extremely deep/large
plain arrays, or opaque host objects is stored normally but is not retained for
undo. Writing such metadata creates the same history barrier, because its
reachable memory cannot be cloned and bounded reliably.

`transact` callbacks must be synchronous. Promise-returning callbacks are
rejected and their captured mutations are rolled back.

Explicit transactions are atomic within the configured history budget. If a
transaction would exceed `maxBytes` or write unsupported metadata, it throws
and restores its starting state. The same oversized mutation performed as a
normal single engine operation succeeds and creates a non-undoable barrier.

## Development Status

### ✅ Completed

- Core type system and interfaces
- Basic engine structure with sheet management
- Cell addressing system (A1 notation)
- Sparse data storage with Map-based implementation
- Copy/paste operations
- Named expressions (storage only)

### 🚧 In Progress

- Formula parser and lexer
- Formula evaluation engine
- Dependency tracking system
- Array formula support

### 📋 Planned

- Function library (SUM, AVERAGE, etc.)
- Array formulas with broadcasting
- Comprehensive error handling
- React hooks for integration
- Performance optimizations

## Running Tests

```bash
bun test
```

## Architecture

FormulaEngine uses a sparse-aware architecture optimized for spreadsheets where most cells are empty:

- **Sheets** store cells in a `Map<string, Cell>` structure
- **Addresses** use zero-based indexing internally, A1 notation externally
- **Formulas** will be parsed into ASTs for efficient evaluation
- **Dependencies** will be tracked in a directed acyclic graph

## License

MIT
