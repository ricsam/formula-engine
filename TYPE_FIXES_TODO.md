# Type Fixes TODO

## Current Status (2024-12-30)

### Type Errors Identified
1. **Test files**: Using object literals instead of proper AST node types
2. **Array results**: Tests expect arrays but `CellValue` type doesn't include arrays

### Temporary Fixes Applied
1. Integration tests use type assertions for array results: `as unknown as number[][]`
2. Dependency graph tests use optional chaining and type assertions

### Recommended Permanent Fixes

#### 1. Update Test Helpers
The test-helpers.ts file has been created with proper factory functions for creating AST nodes. All unit tests should be updated to use these helpers instead of object literals.

#### 2. Consider Extended Result Type
For testing purposes, consider creating an extended result type:
```typescript
type EvaluatorTestResult = CellValue | CellValue[][];
```

#### 3. Complete Test Updates Needed
- `tests/unit/evaluator/evaluator.test.ts`: Replace all object literals with helper functions
- All array result assertions should use proper type handling

### Why These Type Errors Don't Affect Functionality
1. The evaluator correctly returns arrays for array operations
2. The engine will handle spilling (not yet implemented)
3. Tests are correctly validating the behavior despite type assertions

### Priority
These are TypeScript compilation errors only - the JavaScript execution is correct. The fixes are important for maintainability but don't affect runtime behavior.