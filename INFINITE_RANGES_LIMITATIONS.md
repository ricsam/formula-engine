# Infinite Ranges: Known Limitations and Solutions

This document outlines the current limitations in the FormulaEngine's infinite range implementation and provides detailed technical solutions for addressing them.

## 1. Circular Reference Detection with Infinite Ranges

### Current Limitation

The FormulaEngine currently cannot detect circular references that involve infinite ranges. For example:

```javascript
// This should result in #CYCLE! errors but currently doesn't
engine.setCellContents('B1', '=SUM(A:A)');  // B1 depends on all of column A
engine.setCellContents('A1', '=B1');        // A1 (part of A:A) depends on B1
```

**Expected behavior**: Both A1 and B1 should show `#CYCLE!`
**Current behavior**: Values are cached and no cycle is detected

### Root Cause

The issue stems from the evaluation architecture:

1. **Context Isolation**: The `getRangeValues` function doesn't receive the evaluation context with the current evaluation stack
2. **Direct Cell Access**: When evaluating ranges, the engine accesses `cell.value` directly instead of going through `getCellValue` which performs cycle detection
3. **Cache Dependencies**: Cells are evaluated once and cached, preventing re-evaluation that would reveal the cycle

### Technical Solution

#### Phase 1: Modify Range Evaluation Context

Update the `EvaluationContext` to include range-aware cycle detection:

```typescript
// In src/evaluator/evaluator.ts
export interface EvaluationContext {
  currentSheet: number;
  currentCell?: SimpleCellAddress;
  namedExpressions: Map<string, NamedExpression>;
  getCellValue: (address: SimpleCellAddress) => CellValue;
  getRangeValues: (range: SimpleCellRange, evaluationStack?: Set<string>) => CellValue[][];  // Add stack parameter
  getFunction: (name: string) => FunctionDefinition | undefined;
  errorHandler: ErrorHandler;
  evaluationStack: Set<string>;
  arrayContext?: ArrayEvaluationContext;
  sheetResolver?: SheetResolver;
}
```

#### Phase 2: Update getRangeValuesInternal

Modify the range evaluation to accept and use the evaluation stack:

```typescript
// In src/core/engine.ts
private getRangeValuesInternal(range: SimpleCellRange, evaluationStack?: Set<string>): CellValue[][] {
  const sheet = this.sheets.get(range.start.sheet);
  if (!sheet) return [[]];
  
  // Check if this is an infinite range
  const isInfiniteColumn = range.end.row === Number.MAX_SAFE_INTEGER;
  const isInfiniteRow = range.end.col === Number.MAX_SAFE_INTEGER;
  
  if (isInfiniteColumn || isInfiniteRow) {
    return this.getInfiniteRangeValues(sheet, range, isInfiniteColumn, isInfiniteRow, evaluationStack);
  }
  
  // Normal range handling with cycle detection
  const result: CellValue[][] = [];
  
  for (let row = range.start.row; row <= range.end.row; row++) {
    const rowValues: CellValue[] = [];
    for (let col = range.start.col; col <= range.end.col; col++) {
      const cellAddress = { sheet: range.start.sheet, row, col };
      
      // Use getCellValueInternal with evaluation stack for cycle detection
      const value = evaluationStack 
        ? this.getCellValueInternal(cellAddress, evaluationStack)
        : this.getCellValue(cellAddress);
        
      rowValues.push(value);
    }
    result.push(rowValues);
  }
  
  return result;
}
```

#### Phase 3: Update Context Creation

Ensure the evaluation context passes the stack through:

```typescript
// In src/core/engine.ts (setCellContents method)
const context: EvaluationContext = {
  currentSheet: address.sheet,
  currentCell: address,
  namedExpressions: this.namedExpressions,
  getCellValue: (addr: SimpleCellAddress) => this.getCellValueInternal(addr, context.evaluationStack),
  getRangeValues: (range: SimpleCellRange) => this.getRangeValuesInternal(range, context.evaluationStack), // Pass stack
  getFunction: (name: string) => functionRegistry.get(name),
  errorHandler: this.errorHandler,
  evaluationStack: new Set<string>(),
  sheetResolver: (sheetName: string) => this.getSheetId(sheetName)
};
```

#### Phase 4: Infinite Range Cycle Detection

Update `getInfiniteRangeValues` to check for cycles:

```typescript
private getInfiniteRangeValues(
  sheet: Sheet,
  range: SimpleCellRange,
  isInfiniteColumn: boolean,
  isInfiniteRow: boolean,
  evaluationStack?: Set<string>
): CellValue[][] {
  // ... existing sparse iteration logic ...
  
  // When getting cell values, use evaluation stack
  for (const [key, cell] of sheet.cells) {
    const address = parseA1Key(key, sheet.id);
    if (address && this.isAddressInRange(address, range)) {
      const value = evaluationStack
        ? this.getCellValueInternal(address, evaluationStack)
        : cell.value;
      
      // ... add to result array ...
    }
  }
}
```

### Implementation Complexity

- **Effort**: Medium-High (2-3 days)
- **Risk**: Medium (affects core evaluation logic)
- **Testing**: Extensive testing required for all formula types

---

## 2. Supporting Column Filtering for Row Ranges (FILTER Function)

### Current Limitation

The `FILTER` function is designed to filter rows, not columns. When given row ranges, it doesn't perform the expected column-wise filtering:

```javascript
// Expected: Filter columns in row 1 based on conditions in row 2
=FILTER(1:1, 2:2="Yes")  // Should return: Apple, Cherry

// Current behavior: Returns all values in row 1 without filtering
```

### Use Cases

#### Scenario 1: Horizontal Data with Row-based Conditions
```
Row 1: Apple    Banana    Cherry    Date
Row 2: Yes      No        Yes       No
Row 3: =FILTER(1:1, 2:2="Yes")
Expected Result: Apple, Cherry
```

#### Scenario 2: Vertical Data with Column-based Conditions  
```
Col A: Apple     Col B: Yes    Col C: =FILTER(A:A, B:B="Yes")
       Banana            No
       Cherry            Yes  
       Date              Yes
       Elderberry        No
Expected Result: Apple, Cherry, Date
```

### Technical Solution

#### Phase 1: Detect Range Orientation

Enhance the `FILTER` function to detect whether it's working with row ranges or column ranges:

```typescript
// In src/functions/array/array-functions.ts
export const FILTER: FunctionDefinition = {
  name: 'FILTER',
  returnsArray: true,
  evaluate: (args: CellValue[], context: EvaluationContext): CellValue => {
    if (args.length < 2) {
      throw new Error('#VALUE!');
    }

    const sourceArray = args[0];
    
    // Detect if this is a row range vs column range
    const isRowRange = Array.isArray(sourceArray) && 
                      sourceArray.length === 1 && 
                      Array.isArray(sourceArray[0]);
                      
    const isColumnRange = Array.isArray(sourceArray) && 
                         sourceArray.length > 1 && 
                         sourceArray.every(row => Array.isArray(row) && row.length === 1);

    if (isRowRange) {
      return filterRowRange(sourceArray, args.slice(1));
    } else if (isColumnRange) {
      return filterColumnRange(sourceArray, args.slice(1));
    } else {
      return filterStandardArray(sourceArray, args.slice(1));
    }
  }
};
```

#### Phase 2: Implement Row Range Filtering

Create specialized logic for filtering columns within row ranges:

```typescript
function filterRowRange(sourceArray: CellValue[][], conditionArrays: CellValue[]): CellValue {
  const sourceRow = sourceArray[0]; // Single row with multiple columns
  const result: CellValue[] = [];
  
  // Validate all condition arrays are also row ranges with same column count
  for (const condArray of conditionArrays) {
    if (!Array.isArray(condArray) || condArray.length !== 1 || 
        !Array.isArray(condArray[0]) || condArray[0].length !== sourceRow.length) {
      throw new Error('#VALUE!');
    }
  }
  
  // Filter columns based on conditions
  for (let colIndex = 0; colIndex < sourceRow.length; colIndex++) {
    let includeColumn = true;
    
    // Check all conditions for this column
    for (const condArray of conditionArrays) {
      const conditionRow = condArray[0] as CellValue[];
      if (!coerceToBoolean(conditionRow[colIndex])) {
        includeColumn = false;
        break;
      }
    }
    
    if (includeColumn) {
      result.push(sourceRow[colIndex]);
    }
  }
  
  // Return as column vector (standard FILTER output format)
  return result.map(value => [value]) as unknown as CellValue;
}
```

#### Phase 3: Enhance Column Range Filtering

Improve the existing column range filtering to handle multiple condition columns:

```typescript
function filterColumnRange(sourceArray: CellValue[][], conditionArrays: CellValue[]): CellValue {
  const result: CellValue[][] = [];
  
  // Validate condition arrays
  for (const condArray of conditionArrays) {
    if (!Array.isArray(condArray) || condArray.length !== sourceArray.length) {
      throw new Error('#VALUE!');
    }
  }
  
  // Filter rows based on conditions
  for (let rowIndex = 0; rowIndex < sourceArray.length; rowIndex++) {
    let includeRow = true;
    
    // Check all conditions for this row
    for (const condArray of conditionArrays) {
      const conditionValue = Array.isArray(condArray[rowIndex]) 
        ? condArray[rowIndex][0] 
        : condArray[rowIndex];
        
      if (!coerceToBoolean(conditionValue)) {
        includeRow = false;
        break;
      }
    }
    
    if (includeRow) {
      result.push(sourceArray[rowIndex]);
    }
  }
  
  return result as unknown as CellValue;
}
```

#### Phase 4: Update Array Shape Detection

Enhance the array dimension analysis to distinguish between row and column ranges:

```typescript
function analyzeArrayOrientation(array: CellValue[][]): {
  isRowRange: boolean;
  isColumnRange: boolean;
  isStandardArray: boolean;
  rows: number;
  cols: number;
} {
  if (!Array.isArray(array) || array.length === 0) {
    return { isRowRange: false, isColumnRange: false, isStandardArray: false, rows: 0, cols: 0 };
  }
  
  const rows = array.length;
  const cols = Array.isArray(array[0]) ? array[0].length : 1;
  
  // Row range: 1 row, multiple columns
  const isRowRange = rows === 1 && cols > 1;
  
  // Column range: multiple rows, 1 column each
  const isColumnRange = rows > 1 && array.every(row => 
    Array.isArray(row) && row.length === 1
  );
  
  // Standard array: multiple rows and columns
  const isStandardArray = !isRowRange && !isColumnRange;
  
  return { isRowRange, isColumnRange, isStandardArray, rows, cols };
}
```

### Implementation Complexity

- **Effort**: Medium (1-2 days)
- **Risk**: Low-Medium (isolated to FILTER function)
- **Testing**: Comprehensive testing needed for both orientations

### Backward Compatibility

The enhanced FILTER function maintains backward compatibility:
- Existing column-based filtering continues to work unchanged
- Standard 2D array filtering remains unaffected
- Only adds new functionality for row range scenarios

---

## Implementation Priority

1. **FILTER Enhancement** (Recommended first)
   - Lower risk and complexity
   - Provides immediate user value
   - Self-contained changes

2. **Circular Reference Detection** (Recommended second)  
   - Higher complexity but important for correctness
   - Requires careful testing of core evaluation logic
   - May uncover other evaluation edge cases

Both features represent significant improvements to the FormulaEngine's Excel compatibility and robustness.