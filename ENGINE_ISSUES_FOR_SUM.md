# Engine Issues Identified from SUM Function Testing

This document summarizes the engine limitations discovered while testing the SUM function. The SUM function implementation itself is correct, but these engine features need to be implemented.

## 🔴 Critical Parser Issues

### 1. Zero-Argument Functions Not Allowed
**Issue**: `=SUM()` throws "Invalid number of arguments for function SUM"
**Location**: Parser grammar/validation 
**Impact**: Prevents proper error handling by functions
**Expected**: Parser should allow zero arguments and let functions handle validation

### 2. Division by Zero Handling
**Issue**: `=1/0` produces `#ERROR!` instead of `Infinity`
**Location**: Division operator evaluation
**Impact**: Breaks mathematical operations with infinity
**Expected**: Follow JavaScript behavior (1/0 = Infinity, -1/0 = -Infinity)

## 🔴 Critical Evaluator Issues

### 3. 3D References Not Implemented
**Issue**: Parser recognizes `3d-range` nodes but evaluator doesn't handle them
**Location**: `FormulaEvaluator.evaluateNode()` - missing case for "3d-range"
**Impact**: `=SUM(Sheet1:Sheet3!A1)` fails with `#ERROR!`
**Expected**: Should sum same cell/range across multiple sheets

**Parser Output**: Console shows the parser correctly identifies:
```javascript
{
  type: "3d-range",
  startSheet: "Sheet1", 
  endSheet: "Sheet3",
  reference: { ... }
}
```

## 🟡 Cross-Sheet Reference Issues

### 4. Cross-Sheet Ranges Not Supported
**Issue**: `Sheet2!B1:B2` syntax not supported
**Location**: Parser - cross-sheet range parsing
**Impact**: Cannot sum ranges from other sheets
**Expected**: Should parse and evaluate cross-sheet ranges

### 5. Cross-Sheet Cell References
**Issue**: Individual cross-sheet cells like `Sheet2!B1` may not work consistently
**Location**: Parser - cross-sheet cell parsing  
**Impact**: Mixed results with cross-sheet operations
**Expected**: Should reliably parse individual cross-sheet cells

## 🟡 Structured Reference Issues

### 6. Table Column References
**Issue**: `DataTable[ColumnName]` syntax not supported
**Location**: Parser - structured reference parsing
**Impact**: Cannot reference table columns by name
**Expected**: Should parse and resolve `TableName[ColumnName]`

### 7. Table Area References  
**Issue**: `DataTable[#All]`, `DataTable[#Data]` not supported
**Location**: Parser - structured reference area parsing
**Impact**: Cannot reference specific table areas
**Expected**: Should support standard table area specifiers

## 🟡 Named Expression Issues

### 8. Named Ranges in Function Calls
**Issue**: Named expressions that resolve to ranges don't work in function calls
**Location**: Named expression resolution in function context
**Impact**: `=SUM(VALUES_A)` where `VALUES_A = A1:A3` fails
**Expected**: Should expand named ranges in function arguments

## Test Results Summary

```
✅ PASSING (5 tests):
- Basic scalar arguments
- Cell references (A1:A3)  
- Dynamic arrays (SEQUENCE)
- SUM used in dynamic array context
- Mixed argument types

❌ FAILING (6 tests):
- Structured references (Table[Column])
- Named expressions (range expansion) 
- Cross-sheet references (Sheet2!B1:B2)
- 3D references (Sheet1:Sheet3!A1)
- Infinity handling (1/0)
- Error handling (zero arguments)

⏭️ SKIPPED (2 tests):
- Tests moved to commented section
```

## Priority Recommendations

**High Priority** (Breaks core functionality):
1. Fix 3D reference evaluation (parser already works)
2. Fix division by zero to produce Infinity
3. Allow zero-argument functions

**Medium Priority** (Excel compatibility):
4. Implement cross-sheet range parsing
5. Implement structured reference parsing  
6. Fix named range expansion in functions

**Low Priority** (Advanced features):
7. Table area references (#All, #Data)
8. Cross-sheet named expressions

## Files to Review

The comprehensive test cases are in `src/functions/math/sum/sum.test.ts` in the commented section at the bottom. Each test clearly identifies:
- The specific engine limitation
- The expected behavior
- The exact location that needs work

You can uncomment tests one by one as you implement the corresponding engine features.
