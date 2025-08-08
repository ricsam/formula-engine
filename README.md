# FormulaEngine

A TypeScript-based spreadsheet formula evaluation library designed for high-performance calculation of formulas across sparse datasets.

## Features

- **Sparse-aware architecture** - Only populated cells consume memory
- **Map-based storage** - O(1) cell access with automatic sparse handling
- **Excel-compatible** addressing (A1 notation)
- **Multi-sheet support** with sheet management
- **Named expressions** with global and sheet-level scoping
- **Undo/redo system** with command pattern
- **Copy/paste operations** with clipboard support
- **TypeScript-first** design with comprehensive type safety

## Installation

```bash
bun install
```

## Quick Start

```typescript
import { FormulaEngine } from 'formula-engine';

// Create a new engine
const engine = FormulaEngine.buildEmpty();

// Add a sheet
const sheetName = engine.addSheet('Sheet1');
const sheetId = engine.getSheetId(sheetName);

// Set cell values
engine.setCellContent({ sheet: sheetId, col: 0, row: 0 }, 42);
engine.setCellContent({ sheet: sheetId, col: 1, row: 0 }, 58);

// Set a formula (evaluation not yet implemented)
engine.setCellContent({ sheet: sheetId, col: 2, row: 0 }, '=A1+B1');

// Get cell value
const value = engine.getCellValue({ sheet: sheetId, col: 0, row: 0 }); // 42

// Set multiple values at once
engine.setCellContent({ sheet: sheetId, col: 0, row: 2 }, [
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9]
]);

// Get range values
const range = {
  start: { sheet: sheetId, col: 0, row: 2 },
  end: { sheet: sheetId, col: 2, row: 4 }
};
const values = engine.getRangeValues(range); // [[1,2,3],[4,5,6],[7,8,9]]
```

## Development Status

### ✅ Completed
- Core type system and interfaces
- Basic engine structure with sheet management
- Cell addressing system (A1 notation)
- Sparse data storage with Map-based implementation
- Copy/paste operations
- Named expressions (storage only)
- Undo/redo infrastructure

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