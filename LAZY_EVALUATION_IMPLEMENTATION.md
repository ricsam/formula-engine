# Lazy Evaluation Implementation for IF and IFERROR

## Overview
Implemented lazy evaluation for the `IF` and `IFERROR` functions to ensure Excel compatibility and prevent unnecessary evaluation of branches that won't be used.

## Changes Made

### 1. IF Function (`src/functions/logical/if/if.ts`)
**Before**: Both `value_if_true` and `value_if_false` arguments were always evaluated, regardless of the logical test result.

**After**: 
- Only evaluates the `value_if_true` argument when the logical test is true
- Only evaluates the `value_if_false` argument when the logical test is false
- For spilled values, both branches are still evaluated (because different cells may take different paths)

### 2. IFERROR Function (`src/functions/logical/iferror/iferror.ts`)
**Before**: Both `value` and `value_if_error` arguments were always evaluated.

**After**:
- Only evaluates the `value_if_error` argument when the value argument results in an error
- If value is not an error, returns it immediately without evaluating the error handler
- For spilled values, both branches are still evaluated (because different cells may have errors or not)

## Benefits

### 1. Excel Compatibility
Excel uses lazy evaluation for conditional functions, so this implementation ensures formulas behave the same way.

### 2. Prevention of Cyclic Dependencies
The key use case is preventing cyclic dependencies that would otherwise be created:

```excel
A1: 1
B1: =IF(A1=1, 100, C1)
C1: =IF(A1=1, B1, 200)
```

When `A1=1`:
- `B1` evaluates to `100` (doesn't evaluate `C1`)
- `C1` evaluates `B1` which is `100` (result: `100`)

Without lazy evaluation, this would create a circular dependency error.

### 3. Error Avoidance
Functions can now use invalid expressions in unevaluated branches:

```excel
=IF(TRUE, "Success", CEILING(5))  // CEILING(5) has missing arg but won't cause error
=IFERROR(42, CEILING(5))          // CEILING(5) has missing arg but won't be evaluated
```

### 4. Performance Optimization
By not evaluating unnecessary branches, the engine performs less computation, especially for complex formulas.

## Testing

Added comprehensive test suites for both functions covering:
- Basic lazy evaluation (unevaluated branches with errors don't cause issues)
- Cyclic dependency prevention
- Nested lazy evaluation
- Conditional error triggering
- Multiple references to same cell

All tests pass (79 tests in logical functions suite).

## Implementation Notes

### Spilled Values Exception
For dynamic arrays (spilled values), both branches must be evaluated because:
1. Different cells in the spill area may take different paths
2. The spill area calculation needs to consider all possible results

Example:
```excel
A1:A3 = [1, 0, 1]
B1 = =IF(A1:A3, "True", "False")
```

Results in:
- `B1 = "True"` (because A1 = 1)
- `B2 = "False"` (because A2 = 0)
- `B3 = "True"` (because A3 = 1)

### Error Propagation
Errors in evaluated branches are properly propagated:
- IF: Returns error from whichever branch is evaluated
- IFERROR: Returns error handler result when value is an error

## Files Modified
- `src/functions/logical/if/if.ts` - Implemented lazy evaluation
- `src/functions/logical/iferror/iferror.ts` - Implemented lazy evaluation
- `src/functions/logical/if/if.test.ts` - Added lazy evaluation test suite
- `src/functions/logical/iferror/iferror.test.ts` - Added lazy evaluation test suite
