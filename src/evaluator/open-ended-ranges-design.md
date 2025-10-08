# Design Specification: Open-Ended Range Support for Aggregate Functions

## Overview

This document describes the design for supporting open-ended (infinite) ranges in aggregate functions like SUM, AVERAGE, MIN, MAX, etc. The key challenge is efficiently evaluating aggregates over ranges like `B10:D` (infinite rows) or `B10:INFINITY` (infinite rows and columns) without iterating over infinite cells.

## Problem Statement

Currently, aggregate functions throw an error when encountering infinite ranges:
- `SUM(B10:D)` - throws "SUM over an infinite row-range is not implemented"
- `SUM(B1:D)` - throws "SUM over an infinite col-range is not implemented"
- `SUM(B10:INFINITY)` - throws "SUM over an infinite end-range is not implemented"

We need to:
1. Sum all defined cells within the open range
2. Handle spilled values that may be finite or infinite
3. Return INFINITY when encountering infinite spilled values that intersect the range
4. Create a reusable pattern for all aggregate functions

## Proposed Solution

### Core Algorithm

1. **Iterate over defined cells**: Since we can't iterate over infinite ranges, we iterate over all defined cells in the sheet and check if they fall within the open range.

2. **Handle direct values**: For each cell in `rawContent` that falls within the range, evaluate it using `runtimeSafeEvaluatedNode`.

3. **Handle spilled values**: 
   - Identify formula cells that might spill into the range (using frontier candidate algorithm from SUM_INFINITE_RANGES.md)
   - For finite spills, calculate the intersection with the range and evaluate cells
   - For infinite spills, return INFINITY

### API Design

```typescript
// Core utility for handling open-ended ranges
interface OpenRangeEvaluator {
  /**
   * Evaluates all cells within an open-ended range and returns their values
   * @param range - The spreadsheet range (may have infinite bounds)
   * @param sheetName - The sheet to evaluate on
   * @param context - Evaluation context
   * @param evaluator - The FormulaEvaluator instance
   * @returns Array of evaluation results or INFINITY if infinite spill detected
   */
  evaluateCellsInRange(
    range: SpreadsheetRange,
    sheetName: string,
    context: EvaluationContext,
    evaluator: FormulaEvaluator
  ): FunctionEvaluationResult[] | CellInfinity;
}

// Helper to check if a cell is within a range
function isCellInRange(
  cell: LocalCellAddress,
  range: SpreadsheetRange
): boolean;

// Helper to get intersection of two ranges
function getRangeIntersection(
  range1: SpreadsheetRange,
  range2: SpreadsheetRange
): SpreadsheetRange | null;

// Helper to identify frontier spill candidates
function getFrontierCandidates(
  range: SpreadsheetRange,
  sheetContent: Map<string, any>,
  sheetName: string
): CellAddress[];
```

### Implementation Details

#### 1. Cell Iteration Strategy

```typescript
function* iterateCellsInOpenRange(
  range: SpreadsheetRange,
  rawContent: Map<string, any>
): Generator<{ address: LocalCellAddress; value: any }> {
  for (const [key, value] of rawContent) {
    const { rowIndex, colIndex } = parseCellReference(key);
    
    // Check if cell is within range bounds
    if (rowIndex < range.start.row || colIndex < range.start.col) continue;
    
    if (range.end.row.type === "number" && rowIndex > range.end.row.value) continue;
    if (range.end.col.type === "number" && colIndex > range.end.col.value) continue;
    
    yield { 
      address: { rowIndex, colIndex }, 
      value 
    };
  }
}
```

#### 2. Spill Handling

```typescript
function handleSpilledValues(
  spillOrigin: CellAddress,
  spillArea: SpreadsheetRange,
  targetRange: SpreadsheetRange,
  context: EvaluationContext,
  evaluate: Function
): FunctionEvaluationResult[] | CellInfinity {
  // Check for infinite spills
  if (spillArea.end.row.type === "infinity" || spillArea.end.col.type === "infinity") {
    // Check if infinite spill intersects with target range
    const intersects = checkRangeIntersection(spillArea, targetRange);
    if (intersects) {
      return { type: "infinity", sign: "positive" };
    }
  }
  
  // Calculate intersection for finite spills
  const intersection = getRangeIntersection(spillArea, targetRange);
  if (!intersection) return [];
  
  // Evaluate cells in the intersection
  const results: FunctionEvaluationResult[] = [];
  // ... iterate over intersection and evaluate
  
  return results;
}
```

#### 3. Frontier Candidate Detection

Based on the algorithm in SUM_INFINITE_RANGES.md:

```typescript
function getFrontierCandidates(
  range: SpreadsheetRange,
  sheetContent: Map<string, any>,
  sheetName: string
): CellAddress[] {
  const candidates = new Set<string>();
  const formulaCells = new Map<string, LocalCellAddress>();
  
  // Identify all formula cells
  for (const [key, value] of sheetContent) {
    if (typeof value === "string" && value.startsWith("=")) {
      const addr = parseCellReference(key);
      formulaCells.set(key, addr);
    }
  }
  
  // Top frontier (for downward spills)
  const colsToCheck = getColumnsInRange(range, sheetContent);
  for (const col of colsToCheck) {
    const nearestAbove = findNearestAboveFormula(col, range.start.row, formulaCells);
    if (nearestAbove) {
      candidates.add(cellAddressToKey(nearestAbove));
    }
  }
  
  // Left frontier (for rightward spills)
  const rowsToCheck = getRowsInRange(range, sheetContent);
  for (const row of rowsToCheck) {
    const nearestLeft = findNearestLeftFormula(row, range.start.col, formulaCells);
    if (nearestLeft) {
      candidates.add(cellAddressToKey(nearestLeft));
    }
  }
  
  return Array.from(candidates).map(key => ({
    ...parseCellReference(key),
    sheetName
  }));
}
```

### Integration with SUM Function

```typescript
export const SUM: FunctionDefinition = {
  name: "SUM",
  evaluate: function (node, context) {
    // ... existing code ...
    
    const parseResult = (result: FunctionEvaluationResult): ... => {
      // ... existing parseResult logic ...
      
      if (result.type === "spilled-values") {
        const range = result.spillArea;
        
        // Handle open-ended ranges
        if (range.end.col.type === "infinity" || range.end.row.type === "infinity") {
          const openRangeEvaluator = new OpenRangeEvaluator(this);
          const cellValues = openRangeEvaluator.evaluateCellsInRange(
            range,
            result.spillOrigin.sheetName,
            context,
            this
          );
          
          if ("type" in cellValues && cellValues.type === "infinity") {
            return cellValues;
          }
          
          let subTotal = 0;
          for (const cellValue of cellValues) {
            const parsed = parseResult(cellValue);
            if (parsed.type === "error" || parsed.type === "infinity") {
              return parsed;
            }
            subTotal += parsed.value;
          }
          
          return { type: "number", value: subTotal };
        }
        
        // ... existing finite range handling ...
      }
    };
    
    // ... rest of the function ...
  }
};
```

## Testing Strategy

### Unit Tests

1. **Basic Open Range Tests**
   ```typescript
   describe("Open-ended range support", () => {
     it("should sum all values in an open row range", () => {
       // Setup: B10:B15 have values, SUM(B10:B)
       // Expected: Sum of B10:B15
     });
     
     it("should sum all values in an open column range", () => {
       // Setup: B10:E10 have values, SUM(B10:10)
       // Expected: Sum of B10:E10
     });
     
     it("should handle empty open ranges", () => {
       // Setup: No values in B10:, SUM(B10:)
       // Expected: 0
     });
   });
   ```

2. **Spill Handling Tests**
   ```typescript
   describe("Spilled values in open ranges", () => {
     it("should include finite spilled values in range", () => {
       // Setup: A10=SEQUENCE(3,3), SUM(B10:)
       // Expected: Sum includes B10:C12 from spill
     });
     
     it("should return INFINITY for infinite spills", () => {
       // Setup: B100=SEQUENCE(INFINITY), SUM(B10:D)
       // Expected: INFINITY
     });
     
     it("should handle partial spill intersections", () => {
       // Setup: A100=SEQUENCE(1,INFINITY), SUM(B10:D)
       // Expected: Sum of B100:D100 only
     });
     
     it("should handle cross-sheet spills", () => {
       // Setup: Sheet2!A10=SEQUENCE(3,3), SUM(Sheet2!B10:INFINITY)
       // Expected: Sum includes B10:C12 from cross-sheet spill
     });
   });
   ```

3. **Frontier Candidate Tests**
   ```typescript
   describe("Frontier candidate detection", () => {
     it("should detect top frontier candidates", () => {
       // Setup: Formula at C8, SUM(B10:D)
       // Expected: C8 evaluated for potential downward spill
     });
     
     it("should detect left frontier candidates", () => {
       // Setup: Formula at A10, SUM(B10:)
       // Expected: A10 evaluated for potential rightward spill
     });
     
     it("should ignore blocked candidates", () => {
       // Setup: Formula at A8, value at B9, SUM(C10:)
       // Expected: A8 not evaluated (blocked by B9)
     });
     
     it("should detect cross-sheet frontier candidates", () => {
       // Setup: Sheet2!C8=SEQUENCE(5,1), SUM(Sheet2!B10:D)
       // Expected: Sheet2!C8 evaluated for potential downward spill
     });
   });
   ```

### Integration Tests

1. **Cross-sheet References**
   - ✅ Basic cross-sheet open ranges (`Sheet2!B10:B`, `Sheet2!B10:10`)
   - ✅ Cross-sheet spills intersecting with open ranges
   - ✅ Cross-sheet infinite spill detection
   - ✅ Cross-sheet frontier candidates (top and left frontiers)
   - ✅ Blocked cross-sheet frontier candidates
   - ✅ Mixed same-sheet and cross-sheet data in formulas
   - ✅ Cross-sheet error propagation
   - ✅ Cross-sheet circular reference detection
   - ✅ References to non-existent sheets
   - ✅ Complex cross-sheet spill chains
   - ✅ Multiple sheets with open ranges

2. **Circular Dependencies**
   - ✅ Test that circular references are detected when evaluating frontier candidates
   - ✅ Ensure proper error propagation across sheets

3. **Performance Tests**
   - ✅ Large sheets with thousands of cells
   - ✅ Multiple nested open ranges
   - ✅ Complex spill patterns

### Edge Cases

1. **Multiple Infinite Spills**
   - Multiple SEQUENCE(INFINITY) formulas in range
   - Overlapping infinite spills

2. **Dynamic Spills**
   - Spills that change size based on other cell values
   - Conditional spills (IF statements returning arrays)

3. **Error Propagation**
   - Cells with errors in open ranges
   - Spilled errors

## Implementation Phases

### Phase 1: Core Infrastructure
1. Implement `OpenRangeEvaluator` class
2. Add helper functions for range operations
3. Update SUM function to use new infrastructure

### Phase 2: Frontier Candidates
1. Implement frontier candidate detection algorithm
2. Add caching for performance
3. Handle transitive dependencies

### Phase 3: Other Aggregate Functions
1. Update AVERAGE to support open ranges
2. Update MIN/MAX to support open ranges
3. Create shared utilities for all aggregate functions

### Phase 4: Optimization
1. Add indexing for faster cell lookups
2. Implement lazy evaluation for large ranges
3. Add result caching for repeated evaluations

## Performance Considerations

1. **Indexing**: Maintain column/row indices for O(1) lookups
2. **Caching**: Cache frontier candidates per range
3. **Lazy Evaluation**: Only evaluate cells when needed
4. **Early Termination**: Stop evaluation when INFINITY is detected

## Future Enhancements

1. **Bi-directional Spills**: Support for upward/leftward spills
2. **Custom Aggregates**: Allow user-defined aggregate functions
3. **Streaming Evaluation**: Process large ranges in chunks
4. **Parallel Evaluation**: Evaluate independent cells in parallel

## Conclusion

This design provides a robust foundation for supporting open-ended ranges in aggregate functions. The key innovation is using the frontier candidate algorithm to efficiently identify which cells need evaluation, combined with smart handling of infinite spills. The modular design allows easy extension to other aggregate functions and future enhancements.
