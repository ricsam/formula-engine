# FormulaEngine Technical Specification

## Overview

**FormulaEngine** is a TypeScript-based spreadsheet formula evaluation library designed for high-performance calculation of formulas across sparse datasets. Unlike traditional spreadsheet engines that assume dense grids, FormulaEngine operates on a **sparse-aware** model where only defined (non-empty) cells consume memory and processing resources.

The engine supports **matrix formulas** with NumPy-style broadcasting, **named expressions** with global and sheet-level scoping, and maintains a comprehensive **dependency graph** for efficient recalculation. It provides both lazy and eager evaluation strategies, making it suitable for applications ranging from lightweight calculators to complex financial modeling systems.

## Core Concepts

### Internal Data Storage

FormulaEngine uses Map-based storage internally for optimal sparse data handling:

```typescript
// Internal sheet representation
interface Sheet {
  id: number;
  name: string;
  cells: Map<string, Cell>; // Key format: "A1", "B2", etc.
  dimensions: SheetDimensions;
}

interface Cell {
  value: CellValue;
  formula?: string;
  type: CellType;
  dependencies?: Set<string>;
  dependents?: Set<string>;
}
```

This Map-based approach provides:

- **O(1) cell access** by address string
- **Automatic sparse handling** - empty cells don't exist in the Map
- **Memory efficiency** - only populated cells consume memory
- **Fast iteration** over populated cells only

### Cells and Values

#### Cell Types

```typescript
type CellType = "FORMULA" | "VALUE" | "ARRAY" | "EMPTY";
```

#### Value Types

```typescript
type CellValueType = "NUMBER" | "STRING" | "BOOLEAN" | "ERROR" | "EMPTY";
type CellValueDetailedType = CellValueType;
type EmptyCellValue = undefined;
type CellValue = number | string | boolean | FormulaError | EmptyCellValue;
```

#### Raw Cell Content

```typescript
type RawCellContent = CellValue | EmptyCellValue;
```

### Sheets and Addressing

Each sheet is identified by a numeric ID and optional name. Cell addresses use zero-based indexing:

```typescript
interface SimpleCellAddress {
  sheet: number;
  col: number;
  row: number;
}

interface SimpleCellRange {
  start: SimpleCellAddress;
  end: SimpleCellAddress;
}
```

### Named Expressions

Named expressions can be scoped globally or to specific sheets:

- **Global scope**: Accessible from any sheet
- **Sheet scope**: Accessible only within the specified sheet
- Names are case-sensitive

```typescript
interface NamedExpression {
  name: string;
  expression: string;
  scope?: number; // undefined = global scope
}
```

### Array Formulas and Broadcasting

FormulaEngine supports NumPy-style broadcasting for array operations:

#### Broadcasting Rules

1. **Scalar + Array**: Scalar is broadcast to all array elements
2. **Compatible Arrays**: Arrays with compatible dimensions are element-wise operated
3. **Auto-expansion**: Smaller arrays are automatically expanded to match larger ones where possible

#### Examples

```javascript
// Scalar broadcasting
=A1:A3 + 10  // Adds 10 to each cell in range

// Array arithmetic with broadcasting
=A1:A3 * B1:B3  // Element-wise multiplication

// Mixed operations
=SUM(A1:A5 * B1:B5) + C1  // Array multiplication then sum, add scalar
```

## Internal Architecture

### Dependency Graph

The engine maintains a directed acyclic graph (DAG) where:

- **Nodes** represent cells, ranges, and named expressions
- **Edges** represent dependencies between nodes
- **Cycles** are detected and result in `#CYCLE!` errors

#### Graph Optimization

- **Range Decomposition**: Large ranges are decomposed into smaller sub-ranges for efficient updates
- **Incremental Updates**: Only affected nodes are recalculated when dependencies change
- **Sparse Representation**: Only populated cells and their dependencies are tracked

### Formula Parsing and AST

Formulas are parsed into Abstract Syntax Trees (ASTs) using a recursive descent parser:

```typescript
interface FormulaAST {
  type: "function" | "reference" | "value" | "operator" | "array";
  value?: any;
  children?: FormulaAST[];
  functionName?: string;
  operator?: string;
}
```

### Evaluation Strategies

#### Lazy Evaluation

- Formulas are evaluated only when their values are requested
- Results are cached until dependencies change
- Optimal for scenarios with infrequent access patterns

#### Eager Evaluation

- Formulas are immediately recalculated when dependencies change
- All affected cells are updated in topological order
- Optimal for real-time applications requiring immediate updates

### Error Handling

FormulaEngine implements comprehensive error handling:

```typescript
type FormulaError =
  | "#DIV/0!" // Division by zero
  | "#N/A" // Value not available
  | "#NAME?" // Invalid name/function
  | "#NUM!" // Invalid number
  | "#REF!" // Invalid reference
  | "#VALUE!" // Wrong argument type
  | "#CYCLE!" // Circular reference
  | "#ERROR!"; // General error
```

Errors propagate through formulas but don't halt evaluation of independent cells.

## API Reference

### Core Data Access

```typescript
// Single cell operations
getCellValue(cellAddress: SimpleCellAddress): CellValue
getCellFormula(cellAddress: SimpleCellAddress): string
getCellSerialized(cellAddress: SimpleCellAddress): RawCellContent

// Sheet-wide operations
getSheetValues(sheetId: number): CellValue[][]
getSheetFormulas(sheetId: number): (string | null)[][]
getSheetSerialized(sheetId: number): RawCellContent[][]

// Multi-sheet operations
getAllSheetsValues(): Record<string, CellValue[][]>
getAllSheetsFormulas(): Record<string, (string | null)[][]>
getAllSheetsSerialized(): Record<string, RawCellContent[][]>
```

### Data Manipulation

```typescript
// Cell content modification - supports both Map and array formats
setCellContent(
  topLeftCornerAddress: SimpleCellAddress,
  cellContents: RawCellContent[][] | RawCellContent | Map<string, RawCellContent>
): ExportedChange[]

// Map-based sheet operations
setSheetContent(sheetId: number, contents: Map<string, RawCellContent>): ExportedChange[]
getSheetContents(sheetId: number): Map<string, CellValue>

// Range operations
getRangeValues(source: SimpleCellRange): CellValue[][]
getRangeFormulas(source: SimpleCellRange): (string | null)[][]
getRangeSerialized(source: SimpleCellRange): RawCellContent[][]

// Clipboard operations
copy(source: SimpleCellRange): CellValue[][]
cut(source: SimpleCellRange): CellValue[][]
paste(targetLeftCorner: SimpleCellAddress): ExportedChange[]
```

### Sheet Management

```typescript
// Sheet lifecycle
addSheet(sheetName?: string): string
removeSheet(sheetId: number): ExportedChange[]
renameSheet(sheetId: number, newName: string): void

// Sheet information
getSheetName(sheetId: number): string
getSheetId(sheetName: string): number
doesSheetExist(sheetName: string): boolean
countSheets(): number
```

### Named Expressions

```typescript
// Named expression management
addNamedExpression(
  expressionName: string,
  expression: RawCellContent,
  scope?: number
): ExportedChange[]

changeNamedExpression(
  expressionName: string,
  newExpression: RawCellContent,
  scope?: number
): ExportedChange[]

removeNamedExpression(
  expressionName: string,
  scope?: number
): ExportedChange[]

// Named expression access
getNamedExpressionValue(expressionName: string, scope?: number): CellValue
getNamedExpressionFormula(expressionName: string, scope?: number): string
listNamedExpressions(scope?: number): string[]
```

### Dependency Analysis

```typescript
// Dependency tracking
getCellDependents(address: SimpleCellAddress | SimpleCellRange): (SimpleCellRange | SimpleCellAddress)[]
getCellPrecedents(address: SimpleCellAddress | SimpleCellRange): (SimpleCellRange | SimpleCellAddress)[]
```

### Evaluation Control

```typescript
// Evaluation management
suspendEvaluation(): void
resumeEvaluation(): ExportedChange[]
isEvaluationSuspended(): boolean

// Formula utilities
normalizeFormula(formulaString: string): string
calculateFormula(formulaString: string, sheetId: number): CellValue
validateFormula(formulaString: string): boolean
```

### Undo/Redo System

```typescript
// History management
undo(): ExportedChange[]
redo(): ExportedChange[]
isThereSomethingToUndo(): boolean
isThereSomethingToRedo(): boolean
clearUndoStack(): void
clearRedoStack(): void
```

## Usage Examples

### Basic Operations with Maps

```typescript
import { FormulaEngine } from "./formula-engine";

// Initialize empty engine
const engine = FormulaEngine.buildEmpty();

// Add a sheet
const sheetId = engine.getSheetId(engine.addSheet("Data"));

// Set values using Map structure (preferred for sparse data)
const cellData = new Map([
  ["A1", 1],
  ["B1", 2],
  ["C1", 3],
  ["A2", "=SUM(A1:C1)"],
]);

engine.setSheetContent(sheetId, cellData);

// Alternative: set individual cells
engine.setCellContent({ sheet: sheetId, col: 0, row: 2 }, "Hello");
engine.setCellContent({ sheet: sheetId, col: 1, row: 2 }, "=A1*2");

// Get computed values
const result = engine.getCellValue({ sheet: sheetId, col: 0, row: 1 }); // 6
const sheet = engine.getSheetContents(sheetId);
// Returns: Map(['A1' => 1, 'B1' => 2, 'C1' => 3, 'A2' => 6, 'A3' => 'Hello', 'B3' => 2])
```

### Working with Sparse Data

```typescript
// Large sparse dataset - only populated cells consume memory
const playersB = new Map([
  ["A1", "7"],
  ["B1", "19"],
  ["A2", "8"],
  ["B2", "31"],
  ["A3", "9"],
  ["B3", "61"],
  ["A4", "10"],
  ["B4", "89"],
  ["A5", "11"],
  ["B5", "107"],
  ["A6", "12"],
  ["B6", "127"],
]);

engine.setSheetContent(sheetId, playersB);

// Efficient sum over sparse range
engine.setCellContent({ sheet: sheetId, col: 2, row: 0 }, "=SUM(A1:A1000)");
// Only processes populated cells A1-A6, ignoring empty A7-A1000
```

### Named Expressions

```typescript
// Global named expression
engine.addNamedExpression("TaxRate", 0.08);
engine.addNamedExpression("Revenue", "=SUM(Data!A1:A100)");

// Sheet-scoped named expression
engine.addNamedExpression("LocalDiscount", 0.05, sheetId);

// Use in formulas
engine.setCellContent({ sheet: sheetId, col: 0, row: 2 }, [
  ["=Revenue * (1 - TaxRate)"],
]);
```

### Array Formulas with Broadcasting

```typescript
// Set up data using Map structure
const arrayData = new Map([
  ["A1", 1],
  ["B1", 2],
  ["C1", 3],
  ["A2", 4],
  ["B2", 5],
  ["C2", 6],
  ["A3", 7],
  ["B3", 8],
  ["C3", 9],
]);

engine.setSheetContent(sheetId, arrayData);

// Array formula with broadcasting
engine.setCellContent(
  { sheet: sheetId, col: 4, row: 0 },
  "=A1:C3 * 2" // Multiplies each cell by 2
);

// Mixed array operations
engine.setCellContent(
  { sheet: sheetId, col: 4, row: 4 },
  "=SUM(A1:A3 * B1:B3) + 100" // Array multiplication then sum
);
```

### Sparse Operations

```typescript
// Large sparse range - only processes populated cells
const sparseData = new Map([
  ["A1000", 42],
  ["A5000", 84],
]);

engine.setSheetContent(sheetId, sparseData);

// SUM over large range - efficiently processes only 2 cells
engine.setCellContent(
  { sheet: sheetId, col: 1, row: 0 },
  "=SUM(A1:A10000)" // Result: 126, processes only 2 populated cells
);
```

### Dependency Analysis

```typescript
// Set up dependencies using Map
const dependencyData = new Map([
  ["A1", "=B1+C1"],
  ["B1", 10],
  ["C1", 20],
]);

engine.setSheetContent(sheetId, dependencyData);

// Analyze dependencies
const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 0, row: 0 });
// Returns: [{sheet: 0, col: 1, row: 0}, {sheet: 0, col: 2, row: 0}]

const dependents = engine.getCellDependents({ sheet: sheetId, col: 1, row: 0 });
// Returns: [{sheet: 0, col: 0, row: 0}]
```

### Error Handling

```typescript
// Division by zero
const errorData = new Map([
  ["A1", "=1/0"],
  ["A2", "=B2"], // Circular reference setup
  ["B2", "=A2"], // Circular reference setup
]);

engine.setSheetContent(sheetId, errorData);

const error = engine.getCellValue({ sheet: sheetId, col: 0, row: 0 }); // '#DIV/0!'
const cycle1 = engine.getCellValue({ sheet: sheetId, col: 0, row: 1 }); // '#CYCLE!'
const cycle2 = engine.getCellValue({ sheet: sheetId, col: 1, row: 1 }); // '#CYCLE!'
```

## React Integration

### Hook Implementation

```typescript
import { useState, useEffect } from "react";
import type { FormulaEngine, CellValue } from "../formula-engine";

export function useSerializedSheet(
  engine: FormulaEngine,
  sheetId: number
): Map<string, CellValue> {
  const [serialized, setSerialized] = useState<Map<string, CellValue>>(
    new Map()
  );

  React.useEffect(() => {
    return engine.onCellsUpdate(sheetId, (events) => {
      setSerialized(engine.getSheetSerialized(sheetId));
    });
  }, [engine, sheetId]);

  return serialized;
}

// Hook for single cell value
export function useCellValue(
  engine: FormulaEngine,
  sheetId: number,
  cellAddress: string
): CellValue {
  const [cellValue, setCellValue] = useState<CellValue>(() => {
    const sheetId = engine.getSheetId(sheetName);
    const address = engine.simpleCellAddressFromString(cellAddress, sheetId);
    return engine.getCellValue(address);
  });

  useEffect(() => {
    const sheetId = engine.getSheetId(sheetName);
    const address = engine.simpleCellAddressFromString(cellAddress, sheetId);

    const updateCell = () => {
      const value = engine.getCellValue(address);
      setCellValue(value);
    };

    return engine.onCellUpdate(address, (changedAddress) => {
      updateCell();
    });
  }, [engine, sheetName, cellAddress]);

  return cellValue;
}
```

## Performance Characteristics

### Memory Usage

- **O(n)** where n = number of populated cells
- No memory allocation for empty cells or ranges
- Efficient storage for large sparse datasets

### Computational Complexity

- **Formula evaluation**: O(d) where d = number of dependencies
- **Dependency updates**: O(log n) for graph traversal
- **Range operations**: O(p) where p = populated cells in range

### Optimization Strategies

- **Incremental recalculation**: Only affected cells are recomputed
- **Range decomposition**: Large ranges split into manageable chunks
- **AST reuse**: Common formula patterns share parsed representations
- **Lazy evaluation**: Computation deferred until values are needed

## Implementation Notes

### Type System

FormulaEngine uses strict TypeScript typing throughout:

- No type coercion between incompatible types
- Explicit error values for type mismatches
- Clear separation between cell types and value types

### Threading Model

- Single-threaded synchronous evaluation
- All operations return immediately with results or throw errors
- Suitable for embedding in UI frameworks with event loops

### Extensibility

- Modular function library supporting custom functions
- Pluggable parsers for different formula syntaxes
- Configurable error handling and evaluation strategies

This specification provides the foundation for implementing a robust, efficient formula engine suitable for modern web applications requiring spreadsheet-like calculation capabilities.
