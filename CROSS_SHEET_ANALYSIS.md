# Cross-Sheet References Implementation

## Summary

I've successfully implemented cross-sheet references in FormulaEngine. Here's what was done:

### Implementation Overview

**✅ What Was Fixed:**

1. **Parser Enhancement**:
   - Added `SheetResolver` type to allow sheet name resolution
   - Modified parser to accept an optional sheet resolver function
   - Updated `parseIdentifier` to handle sheet references (e.g., `Sheet1!A1`)
   - Added support for quoted sheet names (e.g., `'My Sheet'!A1`)
   - Fixed cross-sheet range parsing (e.g., `Sheet1!A1:A3`)

2. **Engine Integration**:
   - Engine now passes its `getSheetId` method as the sheet resolver
   - All `parseFormula` calls include the sheet resolver
   - `EvaluationContext` now includes the sheet resolver

3. **Error Handling**:
   - Sheet not found errors properly return `#REF!`
   - Sheet deletion triggers re-evaluation of dependent formulas
   - Formulas referencing deleted sheets return `#REF!`

### Key Changes Made

1. **`src/parser/parser.ts`**:
   ```typescript
   export type SheetResolver = (sheetName: string) => number;
   
   // Parser now accepts sheet resolver
   constructor(tokens: Token[], contextSheetId: number = 0, sheetResolver?: SheetResolver)
   
   // Sheet name resolution with error handling
   private getSheetId(sheetName: string): number {
     if (this.sheetResolver) {
       const sheetId = this.sheetResolver(sheetName);
       if (sheetId === -1) {
         throw new ParseError(`Sheet '${sheetName}' not found`);
       }
       return sheetId;
     }
     return this.contextSheetId;
   }
   ```

2. **`src/core/engine.ts`**:
   ```typescript
   // Pass sheet resolver when parsing
   const ast = parseFormula(formula, address.sheet, (sheetName) => this.getSheetId(sheetName));
   
   // Enhanced removeSheet to re-evaluate dependent formulas
   removeSheet(sheetId: number): ExportedChange[] {
     // ... delete sheet ...
     // Force re-evaluation of formulas referencing deleted sheet
     for (const [remainingSheetId, remainingSheet] of this.sheets) {
       const formulaCells = getFormulaCells(remainingSheet);
       // Check and re-evaluate cells that reference the deleted sheet
     }
   }
   ```

3. **`src/evaluator/evaluator.ts`**:
   ```typescript
   // EvaluationContext now includes sheet resolver
   export interface EvaluationContext {
     // ... existing fields ...
     sheetResolver?: SheetResolver;
   }
   ```

### Test Results

All 17 cross-sheet reference tests now pass:
- ✅ Simple cross-sheet references (`Sheet1!A1`)
- ✅ Quoted sheet names (`'My Sheet'!A1`)
- ✅ Cross-sheet ranges (`Sheet1!A1:A3`)
- ✅ Absolute references (`Sheet1!$A$1`)
- ✅ Cross-sheet arithmetic operations
- ✅ Functions with cross-sheet arguments
- ✅ Dependency tracking across sheets
- ✅ Error handling for non-existent sheets
- ✅ Sheet deletion updates references to `#REF!`
- ✅ Complex scenarios (VLOOKUP, IF, FILTER across sheets)

### Technical Details

**Parser Token Flow**:
1. Lexer tokenizes `Sheet1!A1` as: `IDENTIFIER(Sheet1)`, `EXCLAMATION(!)`, `IDENTIFIER(A1)`
2. Parser's `parseIdentifier` detects the exclamation mark
3. Sheet name is extracted and resolved via the sheet resolver
4. Cell reference is parsed with the resolved sheet ID

**Cross-Sheet Range Handling**:
- For ranges like `Sheet1!A1:A3`, the sheet name is prepended to both start and end references
- This ensures consistent sheet resolution across the entire range

**Error Propagation**:
- `ParseError` with "not found" message → `#REF!` error
- Sheet deletion invalidates cached formula values and forces re-evaluation
- Circular references across sheets properly return `#CYCLE!`

### Files Modified

1. **`src/parser/parser.ts`**:
   - Added `SheetResolver` type
   - Updated constructor and static parse method
   - Modified `parseIdentifier` to handle sheet references
   - Updated `parseCellOrRangeWithSheet` for range support
   - Enhanced `getSheetId` with proper error handling

2. **`src/core/engine.ts`**:
   - Imported `SheetResolver` and `ParseError` types
   - Updated all `parseFormula` calls to include sheet resolver
   - Enhanced `removeSheet` to re-evaluate dependent formulas
   - Added sheet resolver to all `EvaluationContext` creations
   - Improved error handling to map ParseError to `#REF!`

3. **`src/evaluator/evaluator.ts`**:
   - Added `sheetResolver` to `EvaluationContext` interface
   - Updated `parseFormula` call to use context's sheet resolver

4. **`src/core/sheet.ts`**:
   - Exported `getFormulaCells` function for finding formula cells

### Conclusion

Cross-sheet references are now fully functional in FormulaEngine, with proper error handling and dependency tracking. The implementation maintains backward compatibility while adding this essential spreadsheet feature. All 573 tests pass, confirming that the implementation is robust and doesn't break existing functionality.