# FormulaEngine Development Plan

## Project Structure

```
formula-engine/
├── src/
│   ├── core/
│   │   ├── engine.ts              # Main FormulaEngine class
│   │   ├── sheet.ts               # Sheet management
│   │   ├── cell.ts                # Cell data structures
│   │   ├── address.ts             # Cell addressing utilities
│   │   └── types.ts               # Core type definitions
│   ├── parser/
│   │   ├── lexer.ts               # Token lexer
│   │   ├── parser.ts              # Formula parser
│   │   ├── ast.ts                 # AST node definitions
│   │   └── grammar.ts             # Grammar rules
│   ├── evaluator/
│   │   ├── evaluator.ts           # Formula evaluation engine
│   │   ├── dependency-graph.ts    # Dependency tracking
│   │   ├── array-evaluator.ts     # Array formula evaluation
│   │   └── error-handler.ts       # Error management
│   ├── functions/
│   │   ├── index.ts               # Function registry
│   │   ├── math/
│   │   │   ├── basic.ts           # +, -, *, /, ^, %
│   │   │   ├── advanced.ts        # SIN, COS, LOG, etc.
│   │   │   └── statistical.ts     # SUM, COUNT, AVERAGE, etc.
│   │   ├── logical/
│   │   │   ├── conditions.ts      # IF, IFS, AND, OR, NOT
│   │   │   └── comparisons.ts     # Comparison operators
│   │   ├── text/
│   │   │   └── string-functions.ts # CONCATENATE, LEN, TRIM, etc.
│   │   ├── lookup/
│   │   │   └── lookup-functions.ts # VLOOKUP, INDEX, MATCH, etc.
│   │   ├── info/
│   │   │   └── info-functions.ts   # ISBLANK, ISERROR, SHEET, etc.
│   │   └── array/
│   │       └── array-functions.ts  # FILTER, ARRAY_CONSTRAIN
│   ├── named-expressions/
│   │   ├── manager.ts             # Named expression management
│   │   └── scope.ts               # Scoping rules
│   ├── operations/
│   │   ├── undo-redo.ts           # Command pattern for undo/redo
│   │   ├── clipboard.ts           # Copy/paste operations
│   │   └── sheet-operations.ts    # Row/column operations
│   ├── utils/
│   │   ├── sparse-map.ts          # Sparse data utilities
│   │   ├── reference-transformer.ts # Reference adjustment
│   │   └── validation.ts          # Input validation
│   └── react/
│       ├── hooks.ts               # React integration hooks
│       └── types.ts               # React-specific types
├── tests/
│   ├── unit/
│   │   ├── core/
│   │   ├── parser/
│   │   ├── evaluator/
│   │   ├── functions/
│   │   └── operations/
│   ├── integration/
│   │   ├── excel-compatibility/
│   │   └── performance/
│   └── fixtures/
│       ├── test-workbooks/
│       └── expected-results/
├── docs/
│   ├── api/
│   ├── functions/
│   └── examples/
├── package.json
├── tsconfig.json
├── bun.lock
├── README.md
└── CHANGELOG.md
```

## Development Phases

### Phase 1: Foundation (Weeks 1-4)
1. **Core Infrastructure**
   - Type system and interfaces
   - Basic engine structure
   - Cell addressing system
   - Sheet management

2. **Parser Foundation**
   - Lexer for tokenization
   - Basic AST structure
   - Simple expression parsing

3. **Basic Evaluation**
   - Simple value evaluation
   - Basic dependency tracking
   - Error handling framework

### Phase 2: Core Functions (Weeks 5-8)
1. **Math Functions**
   - Arithmetic operators
   - Basic math functions
   - Statistical functions

2. **Logical Functions**
   - Conditional logic
   - Boolean operations
   - Comparison operators

3. **Array Support**
   - Array formula evaluation
   - Broadcasting logic
   - Spilling behavior

### Phase 3: Advanced Features (Weeks 9-12)
1. **Lookup Functions**
   - VLOOKUP, HLOOKUP
   - INDEX, MATCH
   - XLOOKUP

2. **Text Functions**
   - String manipulation
   - Text processing

3. **Named Expressions**
   - Global and local scoping
   - Circular reference prevention

### Phase 4: Operations & Integration (Weeks 13-16)
1. **Sheet Operations**
   - Copy/paste mechanics
   - Row/column operations
   - Undo/redo system

2. **React Integration**
   - React hooks
   - Event system
   - Performance optimization

3. **Testing & Documentation**
   - Comprehensive test suite
   - API documentation
   - Performance benchmarks

## File Creation Checklist

### Core Files

#### `src/core/types.ts` - ✅ DONE (2024-12-30)
```typescript
export interface SimpleCellAddress {
  sheet: number;
  col: number;
  row: number;
}

export interface SimpleCellRange {
  start: SimpleCellAddress;
  end: SimpleCellAddress;
}

export type CellValue = number | string | boolean | FormulaError | undefined;
export type RawCellContent = CellValue;
export type CellType = 'FORMULA' | 'VALUE' | 'ARRAY' | 'EMPTY';
export type CellValueType = 'NUMBER' | 'STRING' | 'BOOLEAN' | 'ERROR' | 'EMPTY';
export type CellValueDetailedType = CellValueType;

export interface FormatInfo {
  // Placeholder for future formatting support
}

export interface BoundingRect {
  minCol: number;
  maxCol: number;
  minRow: number;
  maxRow: number;
  width: number;   // maxCol - minCol + 1
  height: number;  // maxRow - minRow + 1
}

export interface ExportedChange {
  address?: SimpleCellAddress;
  oldValue?: CellValue;
  newValue?: CellValue;
  type: 'cell-change' | 'sheet-change' | 'structure-change';
}

export interface NamedExpression {
  name: string;
  expression: string;
  scope?: number;
}

export interface SerializedNamedExpression extends NamedExpression {
  id: string;
}

export interface NamedExpressionOptions {
  overwrite?: boolean;
}

export type FormulaError = 
  | '#DIV/0!'
  | '#N/A'
  | '#NAME?'
  | '#NUM!'
  | '#REF!'
  | '#VALUE!'
  | '#CYCLE!'
  | '#ERROR!';
```

#### `src/core/engine.ts` - ✅ DONE (2024-12-30)
Main FormulaEngine class implementing all API methods:
- Sheet management methods
- Cell access methods
- Named expression methods
- Evaluation control methods
- Undo/redo methods

#### `src/core/sheet.ts` - ✅ DONE (2024-12-30)
Sheet data structure and operations:
- Map-based cell storage
- Bounding rectangle calculation
- Cell enumeration utilities
- Used range tracking

#### `src/core/cell.ts` - ✅ DONE (2024-12-30)
Cell data structures and utilities:
- Cell value storage
- Type detection
- Formula storage

#### `src/core/address.ts` - ✅ DONE (2024-12-30)
Cell addressing utilities:
- Address parsing (A1 notation)
- Address validation
- Range operations
- Address arithmetic

### Parser Files

#### `src/parser/lexer.ts` - ✅ TODO
Token lexer for formula parsing:
- Token types (NUMBER, STRING, FUNCTION, OPERATOR, etc.)
- Tokenization logic
- Error handling

#### `src/parser/ast.ts` - ✅ TODO
AST node definitions:
- Node types for values, functions, operators
- AST traversal utilities
- AST optimization

#### `src/parser/parser.ts` - ✅ TODO
Recursive descent parser:
- Expression parsing
- Function call parsing
- Reference parsing
- Error recovery

#### `src/parser/grammar.ts` - ✅ TODO
Grammar rules and precedence:
- Operator precedence table
- Grammar production rules
- Parsing utilities

### Evaluator Files

#### `src/evaluator/evaluator.ts` - ✅ TODO
Core evaluation engine:
- AST evaluation
- Context management
- Value coercion rules
- Evaluation strategies (lazy/eager)

#### `src/evaluator/dependency-graph.ts` - ✅ TODO
Dependency tracking system:
- Graph data structure
- Cycle detection
- Topological sorting
- Incremental updates

#### `src/evaluator/array-evaluator.ts` - ✅ TODO
Array formula evaluation:
- Broadcasting logic
- Spilling behavior
- Array operation optimization

#### `src/evaluator/error-handler.ts` - ✅ TODO
Error management:
- Error propagation
- Error recovery
- Error formatting

## Function Implementation Checklist

### Math Functions

#### `src/functions/math/basic.ts` - ✅ TODO
- ✅ FE.ADD(a, b) - Addition operator
- ✅ FE.MINUS(a, b) - Subtraction operator  
- ✅ FE.MULTIPLY(a, b) - Multiplication operator
- ✅ FE.DIVIDE(a, b) - Division operator
- ✅ FE.POW(a, b) - Exponentiation operator
- ✅ FE.UMINUS(a) - Unary minus
- ✅ FE.UPLUS(a) - Unary plus
- ✅ FE.UNARY_PERCENT(a) - Percent operator

#### `src/functions/math/advanced.ts` - ✅ TODO
- ✅ ABS(number) - Absolute value
- ✅ SIGN(number) - Sign function
- ✅ SQRT(number) - Square root
- ✅ POWER(base, exponent) - Power function
- ✅ EXP(number) - Exponential function
- ✅ LN(number) - Natural logarithm
- ✅ LOG(number, base) - Logarithm with base
- ✅ LOG10(number) - Base-10 logarithm
- ✅ SIN(number) - Sine function
- ✅ COS(number) - Cosine function  
- ✅ TAN(number) - Tangent function
- ✅ ASIN(number) - Arcsine function
- ✅ ACOS(number) - Arccosine function
- ✅ ATAN(number) - Arctangent function
- ✅ ATAN2(x, y) - Two-argument arctangent
- ✅ DEGREES(radians) - Convert radians to degrees
- ✅ RADIANS(degrees) - Convert degrees to radians
- ✅ PI() - Pi constant
- ✅ ROUND(number, digits) - Round to specified digits
- ✅ ROUNDUP(number, digits) - Round up
- ✅ ROUNDDOWN(number, digits) - Round down
- ✅ CEILING(number, significance) - Round up to multiple
- ✅ FLOOR(number, significance) - Round down to multiple
- ✅ INT(number) - Integer part
- ✅ TRUNC(number, digits) - Truncate
- ✅ MOD(dividend, divisor) - Modulo operation
- ✅ EVEN(number) - Round to even integer
- ✅ ODD(number) - Round to odd integer
- ✅ FACT(number) - Factorial
- ✅ DECIMAL(text, radix) - Convert from base to decimal

#### `src/functions/math/statistical.ts` - ✅ TODO
- ✅ SUM(...numbers) - Sum of values
- ✅ PRODUCT(...numbers) - Product of values
- ✅ COUNT(...values) - Count of numbers
- ✅ COUNTBLANK(range) - Count blank cells
- ✅ COUNTIF(range, criteria) - Count with condition
- ✅ SUMIF(range, criteria, sumRange) - Sum with condition
- ✅ SUMIFS(sumRange, criteriaRange1, criteria1, ...) - Sum with multiple conditions
- ✅ MAX(...numbers) - Maximum value
- ✅ MIN(...numbers) - Minimum value
- ✅ MEDIAN(...numbers) - Median value
- ✅ STDEV(...numbers) - Standard deviation
- ✅ VAR(...numbers) - Variance
- ✅ COVAR(data1, data2) - Covariance
- ✅ GAMMA(number) - Gamma function

### Logical Functions

#### `src/functions/logical/conditions.ts` - ✅ TODO
- ✅ IF(test, thenValue, elseValue) - Conditional logic
- ✅ IFS(condition1, value1, condition2, value2, ...) - Multiple conditions
- ✅ SWITCH(expression, value1, result1, value2, result2, ...) - Switch statement
- ✅ AND(...logicals) - Logical AND
- ✅ OR(...logicals) - Logical OR
- ✅ NOT(logical) - Logical NOT
- ✅ XOR(...logicals) - Exclusive OR
- ✅ TRUE() - True constant
- ✅ FALSE() - False constant
- ✅ IFERROR(value, errorValue) - Error handling
- ✅ IFNA(value, naValue) - N/A error handling

#### `src/functions/logical/comparisons.ts` - ✅ TODO
- ✅ FE.EQ(a, b) - Equality comparison
- ✅ FE.NE(a, b) - Inequality comparison
- ✅ FE.LT(a, b) - Less than comparison
- ✅ FE.LTE(a, b) - Less than or equal comparison
- ✅ FE.GT(a, b) - Greater than comparison
- ✅ FE.GTE(a, b) - Greater than or equal comparison

### Text Functions

#### `src/functions/text/string-functions.ts` - ✅ TODO
- ✅ CONCATENATE(...texts) - Concatenate strings
- ✅ FE.CONCAT(text1, text2) - Binary concatenation
- ✅ LEN(text) - String length
- ✅ TRIM(text) - Remove extra spaces
- ✅ UPPER(text) - Convert to uppercase
- ✅ LOWER(text) - Convert to lowercase
- ✅ EXACT(text1, text2) - Exact string comparison
- ✅ TEXT(number, format) - Format number as text

### Lookup Functions

#### `src/functions/lookup/lookup-functions.ts` - ✅ TODO
- ✅ VLOOKUP(searchValue, tableArray, colIndex, exactMatch) - Vertical lookup
- ✅ HLOOKUP(searchValue, tableArray, rowIndex, exactMatch) - Horizontal lookup  
- ✅ INDEX(range, row, col) - Get value by position
- ✅ MATCH(searchValue, array, matchType) - Find position
- ✅ XLOOKUP(lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode) - Advanced lookup
- ✅ CHOOSE(index, ...values) - Choose value by index
- ✅ OFFSET(reference, rows, cols, height, width) - Offset reference
- ✅ COLUMN(reference) - Get column number
- ✅ COLUMNS(array) - Count columns
- ✅ ROW(reference) - Get row number
- ✅ ROWS(array) - Count rows
- ✅ ADDRESS(row, col, absMode, a1Style, sheet) - Create address string
- ✅ FORMULATEXT(reference) - Get formula as text

### Info Functions

#### `src/functions/info/info-functions.ts` - ✅ TODO
- ✅ ISBLANK(value) - Test for blank
- ✅ ISERROR(value) - Test for error
- ✅ ISERR(value) - Test for error (except N/A)
- ✅ ISNA(value) - Test for N/A error
- ✅ ISNUMBER(value) - Test for number
- ✅ ISTEXT(value) - Test for text
- ✅ ISLOGICAL(value) - Test for logical
- ✅ ISNONTEXT(value) - Test for non-text
- ✅ ISFORMULA(reference) - Test if cell has formula
- ✅ ISEVEN(number) - Test for even number
- ✅ ISODD(number) - Test for odd number
- ✅ ISBINARY(value) - Test for binary number
- ✅ ISREF(value) - Test for reference error
- ✅ SHEET(reference) - Get sheet number
- ✅ SHEETS(reference) - Count sheets
- ✅ NA() - Return N/A error

### Array Functions

#### `src/functions/array/array-functions.ts` - ✅ TODO
- ✅ FILTER(sourceArray, ...boolArrays) - Filter array by conditions
- ✅ ARRAY_CONSTRAIN(array, height, width) - Constrain array size

## API Method Implementation Checklist

### Core Data Access - ✅ TODO
- ✅ getCellValue(cellAddress: SimpleCellAddress): CellValue
- ✅ getCellFormula(cellAddress: SimpleCellAddress): string
- ✅ getCellSerialized(cellAddress: SimpleCellAddress): RawCellContent
- ✅ getSheetValues(sheetId: number): Map<string, CellValue>
- ✅ getSheetFormulas(sheetId: number): Map<string, string>
- ✅ getSheetSerialized(sheetId: number): Map<string, RawCellContent>
- ✅ getSheetBoundingRect(sheetId: number): BoundingRect
- ✅ getAllSheetsBoundingRects(): Record<string, BoundingRect>
- ✅ getAllSheetsValues(): Record<string, Map<string, CellValue>>
- ✅ getAllSheetsFormulas(): Record<string, Map<string, string>>
- ✅ getAllSheetsSerialized(): Record<string, Map<string, RawCellContent>>

### Data Manipulation - ✅ TODO
- ✅ setCellContents(topLeftCornerAddress: SimpleCellAddress, cellContents: RawCellContent[][] | RawCellContent): ExportedChange[]
- ✅ setSheetContents(sheetId: number, contents: Map<string, RawCellContent>): ExportedChange[]
- ✅ getSheetContents(sheetId: number): Map<string, CellValue>
- ✅ getRangeValues(source: SimpleCellRange): CellValue[][]
- ✅ getRangeFormulas(source: SimpleCellRange): (string | undefined)[][]
- ✅ getRangeSerialized(source: SimpleCellRange): RawCellContent[][]
- ✅ getFillRangeData(source: SimpleCellRange, target: SimpleCellRange, offsetsFromTarget: boolean): RawCellContent[][]

### Sheet Management - ✅ TODO
- ✅ isItPossibleToAddSheet(sheetName: string): boolean
- ✅ addSheet(sheetName?: string): string
- ✅ isItPossibleToRemoveSheet(sheetId: number): boolean
- ✅ removeSheet(sheetId: number): ExportedChange[]
- ✅ isItPossibleToClearSheet(sheetId: number): boolean
- ✅ clearSheet(sheetId: number): ExportedChange[]
- ✅ isItPossibleToReplaceSheetContent(sheetId: number, values: RawCellContent[][]): boolean
- ✅ setSheetContent(sheetId: number, values: RawCellContent[][]): ExportedChange[]
- ✅ getSheetName(sheetId: number): string
- ✅ getSheetNames(): string[]
- ✅ getSheetId(sheetName: string): number
- ✅ doesSheetExist(sheetName: string): boolean
- ✅ countSheets(): number
- ✅ isItPossibleToRenameSheet(sheetId: number, newName: string): boolean
- ✅ renameSheet(sheetId: number, newName: string): void

### Operations - ✅ TODO
- ✅ removeRows(sheetId: number, ...indexes: number[]): ExportedChange[]
- ✅ removeColumns(sheetId: number, ...indexes: number[]): ExportedChange[]
- ✅ isItPossibleToMoveCells(source: SimpleCellRange, destinationLeftCorner: SimpleCellAddress): boolean
- ✅ moveCells(source: SimpleCellRange, destinationLeftCorner: SimpleCellAddress): ExportedChange[]
- ✅ moveRows(sheetId: number, startRow: number, numberOfRows: number, targetRow: number): ExportedChange[]
- ✅ isItPossibleToMoveColumns(sheetId: number, startColumn: number, numberOfColumns: number, targetColumn: number): boolean
- ✅ moveColumns(sheetId: number, startColumn: number, numberOfColumns: number, targetColumn: number): ExportedChange[]
- ✅ copy(source: SimpleCellRange): CellValue[][]
- ✅ cut(source: SimpleCellRange): CellValue[][]
- ✅ paste(targetLeftCorner: SimpleCellAddress): ExportedChange[]
- ✅ isClipboardEmpty(): boolean
- ✅ clearClipboard(): void

### Address Utilities - ✅ TODO
- ✅ simpleCellAddressFromString(cellAddress: string, contextSheetId: number): SimpleCellAddress
- ✅ simpleCellRangeFromString(cellRange: string, contextSheetId: number): SimpleCellRange
- ✅ simpleCellRangeToString(cellRange: SimpleCellRange, optionsOrContextSheetId: { includeSheetName?: boolean } | number): string

### Dependency Analysis - ✅ TODO
- ✅ getCellDependents(address: SimpleCellAddress | SimpleCellRange): (SimpleCellRange | SimpleCellAddress)[]
- ✅ getCellPrecedents(address: SimpleCellAddress | SimpleCellRange): (SimpleCellRange | SimpleCellAddress)[]

### Cell Information - ✅ TODO
- ✅ getCellType(cellAddress: SimpleCellAddress): CellType
- ✅ doesCellHaveSimpleValue(cellAddress: SimpleCellAddress): boolean
- ✅ doesCellHaveFormula(cellAddress: SimpleCellAddress): boolean
- ✅ isCellEmpty(cellAddress: SimpleCellAddress): boolean
- ✅ isCellPartOfArray(cellAddress: SimpleCellAddress): boolean
- ✅ getCellValueType(cellAddress: SimpleCellAddress): CellValueType
- ✅ getCellValueDetailedType(cellAddress: SimpleCellAddress): CellValueDetailedType
- ✅ getCellValueFormat(cellAddress: SimpleCellAddress): FormatInfo

### Evaluation Control - ✅ TODO
- ✅ suspendEvaluation(): void
- ✅ resumeEvaluation(): ExportedChange[]
- ✅ isEvaluationSuspended(): boolean

### Named Expressions - ✅ TODO
- ✅ isItPossibleToAddNamedExpression(expressionName: string, expression: RawCellContent, scope?: number): boolean
- ✅ addNamedExpression(expressionName: string, expression: RawCellContent, scope?: number, options?: NamedExpressionOptions): ExportedChange[]
- ✅ getNamedExpressionValue(expressionName: string, scope?: number): CellValue
- ✅ getNamedExpressionFormula(expressionName: string, scope?: number): string
- ✅ getNamedExpression(expressionName: string, scope?: number): NamedExpression
- ✅ isItPossibleToChangeNamedExpression(expressionName: string, newExpression: RawCellContent, scope?: number): boolean
- ✅ changeNamedExpression(expressionName: string, newExpression: RawCellContent, scope?: number, options?: NamedExpressionOptions): ExportedChange[]
- ✅ isItPossibleToRemoveNamedExpression(expressionName: string, scope?: number): boolean
- ✅ removeNamedExpression(expressionName: string, scope?: number): ExportedChange[]
- ✅ listNamedExpressions(scope?: number): string[]
- ✅ getAllNamedExpressionsSerialized(): SerializedNamedExpression[]

### Formula Utilities - ✅ TODO
- ✅ normalizeFormula(formulaString: string): string
- ✅ calculateFormula(formulaString: string, sheetId: number): CellValue
- ✅ getNamedExpressionsFromFormula(formulaString: string): string[]
- ✅ validateFormula(formulaString: string): boolean

### Undo/Redo System - ✅ TODO
- ✅ undo(): ExportedChange[]
- ✅ redo(): ExportedChange[]
- ✅ isThereSomethingToUndo(): boolean
- ✅ isThereSomethingToRedo(): boolean
- ✅ clearRedoStack(): void
- ✅ clearUndoStack(): void

### React Integration - ✅ TODO
- ✅ useSpreadsheet(engine: FormulaEngine, sheetName: string): Map<string, CellValue>
- ✅ useCell(engine: FormulaEngine, sheetName: string, cellAddress: string): CellValue
- ✅ useSpreadsheetRange(engine: FormulaEngine, sheetName: string, range: string): Map<string, CellValue>

## LLM Coding Instructions

### Getting Started
1. **Initialize Project Structure**: Create the folder structure as specified above
2. **Start with Core Types**: Implement `src/core/types.ts` first as foundation
3. **Follow Dependency Order**: Implement files in dependency order (types → core → parser → evaluator → functions)

### Development Workflow
1. **Pick a File**: Choose any file marked with ✅ TODO
2. **Read Dependencies**: Understand what types/interfaces the file needs
3. **Implement**: Write the complete implementation
4. **Test**: Create corresponding test file
5. **Mark Complete**: Change ✅ TODO to ✅ DONE in this plan
6. **Update Exports**: Ensure file is properly exported from index files

### Implementation Guidelines

#### For Core Files:
```typescript
// Always start with proper imports and exports
import { ... } from './types';

export class ClassName {
  // Use private fields with TypeScript conventions
  private readonly data = new Map<string, CellValue>();
  
  // Implement all methods from API specification
  public methodName(param: Type): ReturnType {
    // Validate inputs
    if (!param) {
      throw new Error('Invalid parameter');
    }
    
    // Implement logic
    // Return appropriate type
  }
}
```

#### For Function Files:
```typescript
import { FunctionDefinition } from '../types';

export const ADD: FunctionDefinition = {
  name: 'FE.ADD',
  evaluate: (a: number, b: number): number => {
    // Validate inputs
    if (typeof a !== 'number' || typeof b !== 'number') {
      throw new Error('#VALUE!');
    }
    
    // Implement logic
    return a + b;
  },
  
  // Include argument validation
  validateArgs: (args: unknown[]): boolean => {
    return args.length === 2 && 
           args.every(arg => typeof arg === 'number');
  }
};
```

#### For Test Files:
```typescript
import { test, expect, describe } from "bun:test";
import { ClassName } from '../src/path/to/file';

describe('ClassName', () => {
  test('should handle basic case', () => {
    const instance = new ClassName();
    const result = instance.method('input');
    expect(result).toBe('expected');
  });
  
  test('should handle edge cases', () => {
    // Test error conditions
    // Test boundary values
    // Test invalid inputs
  });
});
```

### Marking Progress
When completing a file:
1. Change `✅ TODO` to `✅ DONE` 
2. Add completion date: `✅ DONE (2024-01-15)`
3. Note any deviations from spec in comments
4. Update the README.md with progress

### Testing Strategy
- **Unit Tests**: Each function and method with `bun test`
- **Integration Tests**: Cross-component functionality  
- **Excel Compatibility**: Compare results with Excel
- **Performance Tests**: Sparse data handling efficiency
- **Bounding Rect Tests**: Verify correct bounds calculation

### Error Handling
- Use TypeScript strict mode
- Validate all inputs
- Return appropriate error types (#VALUE!, #REF!, etc.)
- Include helpful error messages for debugging

### Documentation
- Add JSDoc comments to all public methods
- Include usage examples in complex functions
- Document performance characteristics
- Maintain API compatibility

This development plan provides a clear roadmap for implementing FormulaEngine systematically, with each component building on the previous ones and comprehensive testing throughout the process.