# Table Syntax Support for FormulaEngine

## Overview

Adding Excel-style structured references and table syntax to FormulaEngine requires extending the core architecture to support named table regions, column references, and special table selectors. This document outlines the API design, implementation strategy, and required refactors.

## Table Structure and Core Types

### New Core Types

```typescript
// src/core/types.ts additions

export interface TableDefinition {
  name: string;
  range: SimpleCellRange;
  sheetId: number;
  // Note: headers are mandatory - first row is always headers
}

export interface SerializedTableDefinition extends TableDefinition {
  id: string;
}

export interface TableReference {
  tableName: string;
  columnName?: string;
  selector?: TableSelector;
  isCurrentRow?: boolean; // for [@Column] syntax
}

export type TableSelector = 
  | '#All'        // Header + Data
  | '#Data'       // Data only (default)
  | '#Headers'    // Header row only
  | '#ThisRow';   // Current row context

export interface StructuredReference {
  type: 'structured-reference';
  table: string;
  column?: string;
  selector?: TableSelector;
  thisRow?: boolean;
}
```

### Extended Cell Address Types

```typescript
// Enhanced addressing to support table contexts
export interface CellAddressContext {
  currentTable?: string;
  currentRow?: number;
  evaluationContext?: 'table-formula' | 'regular-formula';
}

export interface TableRegion {
  tableId: string;
  regionType: TableSelector;
  range: SimpleCellRange;
}
```

## API Design

### Table Management API

```typescript
// src/core/engine.ts additions

export class FormulaEngine {
  // Private storage (following named expression pattern)
  private tables: Map<string, TableDefinition> = new Map();

  // Table creation and management (following named expression pattern)
  isItPossibleToAddTable(
    tableName: string,
    range: SimpleCellRange,
    sheetId: number
  ): boolean;

  addTable(
    tableName: string,
    range: SimpleCellRange,
    sheetId: number,
  ): ExportedChange[];
  
  isItPossibleToChangeTable(
    tableName: string,
    newRange: SimpleCellRange,
    sheetId: number
  ): boolean;

  changeTable(
    tableName: string,
    newRange: SimpleCellRange,
    sheetId: number,
  ): ExportedChange[];

  isItPossibleToRemoveTable(tableName: string, sheetId: number): boolean;
  
  removeTable(tableName: string, sheetId: number): ExportedChange[];
  
  // Table information (following named expression pattern)
  getTable(tableName: string, sheetId: number): TableDefinition | undefined;
  getTableByName(tableName: string, sheetId?: number): TableDefinition | undefined;
  listTables(sheetId?: number): string[];
  getAllTablesSerialized(): SerializedTableDefinition[];
  doesTableExist(tableName: string, sheetId?: number): boolean;
  
  // Table data access
  getTableData(tableName: string, sheetId: number, selector?: TableSelector): CellValue[][];
  getTableColumn(tableName: string, sheetId: number, columnName: string): CellValue[];
  
  // Get column names from table header row (first row of table range)
  getTableColumnNames(tableName: string, sheetId: number): string[];
  
  // Validation
  isAddressInTable(address: SimpleCellAddress): TableDefinition | undefined;
  getTableRegion(tableName: string, sheetId: number, selector: TableSelector): SimpleCellRange;
  validateTableRange(range: SimpleCellRange): boolean;

  // Private helpers (following named expression pattern)
  private getTableKey(name: string, sheetId: number): string;
}
```

### Structured Reference Resolution

```typescript
// src/evaluator/table-resolver.ts

export class TableReferenceResolver {
  resolveStructuredReference(
    reference: StructuredReference,
    context: CellAddressContext
  ): SimpleCellRange | SimpleCellAddress;
  
  resolveTableColumn(
    tableName: string,
    columnName: string,
    selector?: TableSelector
  ): SimpleCellRange;
  
  resolveCurrentRowReference(
    tableName: string,
    columnName: string,
    currentRow: number
  ): SimpleCellAddress;
  
  validateStructuredReference(reference: StructuredReference): boolean;
  
  getTableColumnNames(tableName: string): string[];
  
  // Get column names from table header row (always first row of table)
  private getColumnNamesFromHeaders(tableId: string): string[];
}
```

## Parser Extensions

### Grammar Updates

```typescript
// src/parser/grammar.ts additions

// New token types for structured references
export enum TokenType {
  // ... existing tokens
  TABLE_REFERENCE = 'TABLE_REFERENCE',     // Table1[Column1]
  CURRENT_ROW = 'CURRENT_ROW',             // [@Column1]
  TABLE_SELECTOR = 'TABLE_SELECTOR',       // #Data, #Headers, etc.
  COLUMN_SPECIFIER = 'COLUMN_SPECIFIER'    // [Column1]
}

// Grammar rules for structured references
const structuredReferenceRules = {
  // Table1[Column1]
  tableColumnReference: () => seq(
    $.IDENTIFIER,      // Table name
    '[',
    optional(seq('[', $.TABLE_SELECTOR, ']', ',')),  // Optional selector
    $.IDENTIFIER,      // Column name
    ']'
  ),
  
  // [@Column1] or Table1[@Column1]
  currentRowReference: () => seq(
    optional(seq($.IDENTIFIER, '[')),  // Optional table name
    '[@',
    $.IDENTIFIER,      // Column name
    ']'
  ),
  
  // Table1[[#Data],[Column1]]
  fullStructuredReference: () => seq(
    $.IDENTIFIER,      // Table name
    '[[',
    $.TABLE_SELECTOR,  // #Data, #Headers, etc.
    ']',
    optional(seq(',', '[', $.IDENTIFIER, ']')),  // Optional column
    ']'
  )
};
```

### AST Node Extensions

```typescript
// src/parser/ast.ts additions

export interface StructuredReferenceNode extends ASTNode {
  type: 'structured-reference';
  tableName: string;
  columnName?: string;
  selector?: TableSelector;
  isCurrentRow: boolean;
}

export interface TableSelectorNode extends ASTNode {
  type: 'table-selector';
  selector: TableSelector;
}
```

## Evaluator Changes

### Context-Aware Evaluation

```typescript
// src/evaluator/evaluator.ts modifications

export interface EvaluationContext {
  // ... existing context
  currentTable?: string;
  currentTableRow?: number;
  tableResolver: TableReferenceResolver;
}

export class Evaluator {
  evaluateStructuredReference(
    node: StructuredReferenceNode,
    context: EvaluationContext
  ): CellValue | CellValue[] {
    const tableResolver = context.tableResolver;
    
    // Handle current row references [@Column]
    if (node.isCurrentRow) {
      if (!context.currentTable || context.currentTableRow === undefined) {
        throw new FormulaError('#REF!', 'Current row reference outside table context');
      }
      
      const address = tableResolver.resolveCurrentRowReference(
        node.tableName || context.currentTable,
        node.columnName!,
        context.currentTableRow
      );
      
      return this.getCellValue(address);
    }
    
    // Handle full table/column references
    const range = tableResolver.resolveStructuredReference(node, {
      currentTable: context.currentTable,
      currentRow: context.currentTableRow
    });
    
    return this.getRangeValues(range);
  }
}
```

### Dependency Graph Updates

```typescript
// src/evaluator/dependency-graph.ts modifications

export class DependencyGraph {
  // Add table node to dependency graph (following named expression pattern)
  addTable(name: string, sheetId: number): string;
  
  // Creates a unique key for a table (following named expression pattern)
  static getTableKey(name: string, sheetId: number): string;
  
  // Remove table node and update dependencies
  removeTable(tableKey: string): void;
  
  // Get all cells dependent on a table
  getTableDependents(tableKey: string): SimpleCellAddress[];
}
```

## Address and Reference System Changes

### Enhanced Reference Transformer

```typescript
// src/utils/reference-transformer.ts modifications

export class ReferenceTransformer {
  transformStructuredReferences(
    formula: string,
    sourceTable?: string,
    targetTable?: string
  ): string;
  
  // Handle table moves/renames
  updateTableReferences(
    formula: string,
    tableRenames: Map<string, string>
  ): string;
  
  // Handle column renames
  updateColumnReferences(
    formula: string,
    tableName: string,
    columnRenames: Map<string, string>
  ): string;
}
```

### Address Utilities Extension

```typescript
// src/core/address.ts additions

export class AddressUtils {
  parseStructuredReference(reference: string): StructuredReference;
  
  structuredReferenceToString(reference: StructuredReference): string;
  
  convertRangeToStructuredReference(
    range: SimpleCellRange,
    tableName: string,
    columnName?: string
  ): string;
  
  isValidTableName(name: string): boolean;
  isValidColumnName(name: string): boolean;
}
```

## Function Library Extensions

### Table-Aware Functions

```typescript
// src/functions/table/table-functions.ts

export const TABLE_FUNCTIONS = {
  // Convert range to structured reference
  STRUCTURED: {
    name: 'STRUCTURED',
    evaluate: (tableName: string, columnName?: string, selector?: string) => {
      // Return structured reference string
    }
  },
  
  // Get table information
  TABLE_INFO: {
    name: 'TABLE_INFO',
    evaluate: (tableName: string, infoType: string) => {
      // Return table metadata (row count, column count, etc.)
    }
  },
  
  // Table column operations
  TABLE_COLUMN: {
    name: 'TABLE_COLUMN',
    evaluate: (tableName: string, columnName: string) => {
      // Return column data
    }
  }
};
```

### Enhanced Existing Functions

```typescript
// Modifications to existing functions to support table syntax

// VLOOKUP with table syntax: =VLOOKUP(A2, Table1, 2, FALSE)
// INDEX with table syntax: =INDEX(Table1[Sales], MATCH(...))
// SUM with table syntax: =SUM(Table1[Sales])
```

## Storage and Serialization

### Table Metadata Storage

```typescript
// src/core/sheet.ts additions

export class Sheet {
  private tables = new Map<string, TableDefinition>();
  
  addTable(table: TableDefinition): void;
  removeTable(tableId: string): void;
  getTable(tableId: string): TableDefinition | undefined;
  getTableByName(name: string): TableDefinition | undefined;
  getTablesInRange(range: SimpleCellRange): TableDefinition[];
  
  // Serialization support
  getTablesData(): TableDefinition[];
  setTablesData(tables: TableDefinition[]): void;
}
```

### Serialization Format

```typescript
// Extended serialization to include table definitions
export interface SerializedSheet {
  // ... existing properties
  tables: TableDefinition[];
}

export interface SerializedWorkbook {
  // ... existing properties
  tables: TableDefinition[];
}
```

## Validation and Error Handling

### Table-Specific Validation

```typescript
// src/utils/validation.ts additions

export class TableValidator {
  validateTableName(name: string): ValidationResult;
  validateColumnName(name: string): ValidationResult;
  validateTableRange(range: SimpleCellRange): ValidationResult;
  validateStructuredReference(reference: string): ValidationResult;
  
  checkTableOverlap(
    newRange: SimpleCellRange,
    existingTables: TableDefinition[]
  ): ValidationResult;
}

export interface ValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}
```

### Enhanced Error Types

```typescript
// Additional error types for table operations
export type TableError = 
  | '#TABLE!'      // Invalid table reference
  | '#COLUMN!'     // Invalid column reference
  | '#SELECTOR!';  // Invalid table selector

export type FormulaError = 
  | '#DIV/0!'
  | '#N/A'
  | '#NAME?'
  | '#NUM!'
  | '#REF!'
  | '#VALUE!'
  | '#CYCLE!'
  | '#ERROR!'
  | '#TABLE!'      // New
  | '#COLUMN!'     // New
  | '#SELECTOR!';  // New
```

## Operation Support

### Copy/Paste with Tables

```typescript
// src/operations/clipboard.ts modifications

export class ClipboardOperations {
  // Handle copying table data with structured references
  copyTableRange(tableId: string, columns: string[]): CellValue[][];
  
  // Paste with automatic table expansion
  pasteToTable(
    tableId: string,
    data: CellValue[][],
    targetColumn?: string
  ): ExportedChange[];
  
  // Transform structured references during copy/paste
  transformTableReferences(
    formula: string,
    sourceContext: TableContext,
    targetContext: TableContext
  ): string;
}
```

### Undo/Redo for Table Operations

```typescript
// src/operations/undo-redo.ts additions

export class TableCommand implements Command {
  execute(): CommandResult;
  undo(): CommandResult;
  redo(): CommandResult;
}

export class CreateTableCommand extends TableCommand {
  constructor(
    private range: SimpleCellRange,
    private options: TableCreationOptions
  ) {}
}

export class RenameTableCommand extends TableCommand {
  constructor(
    private tableId: string,
    private oldName: string,
    private newName: string
  ) {}
}
```

## React Integration

### Table-Aware Hooks

```typescript
// src/react/hooks.ts additions

export function useTable(
  engine: FormulaEngine,
  tableName: string
): {
  table: TableDefinition | undefined;
  data: Map<string, CellValue>;
  columnNames: string[];
} {
  // React hook for table data
  // columnNames are always read from header row (first row of table)
}

export function useTableColumn(
  engine: FormulaEngine,
  tableName: string,
  columnName: string
): CellValue[] {
  // React hook for specific table column
}

export function useCurrentRowContext(
  engine: FormulaEngine,
  address: SimpleCellAddress
): {
  tableName?: string;
  rowIndex?: number;
  columnValues: Record<string, CellValue>;
} {
  // Hook to get current row context for [@Column] references
}
```

## Implementation Phases

### Phase 1: Core Infrastructure (4-6 weeks)
1. **Type System Extensions**
   - Add table-related types to `src/core/types.ts`
   - Extend addressing system for table contexts

2. **Basic Table Management**
   - Implement table creation, deletion, renaming
   - Add table storage to Sheet class
   - Basic validation for table names and ranges

3. **Parser Extensions**
   - Add tokens for structured references
   - Implement basic grammar rules
   - AST node extensions

### Phase 2: Reference Resolution (3-4 weeks)
1. **Table Reference Resolver**
   - Implement structured reference parsing
   - Add table column resolution
   - Current row context handling

2. **Evaluator Integration**
   - Context-aware evaluation
   - Structured reference evaluation
   - Error handling for invalid references

3. **Dependency Graph Updates**
   - Table-level dependency tracking
   - Table resize dependency updates

### Phase 3: Advanced Features (4-5 weeks)
1. **Copy/Paste Operations**
   - Table-aware clipboard operations
   - Reference transformation during copy/paste
   - Automatic table expansion

2. **Function Library Extensions**
   - Table-aware existing functions
   - New table-specific functions
   - Enhanced VLOOKUP/INDEX with tables

3. **Undo/Redo Support**
   - Table operation commands
   - Complex table modification rollback

### Phase 4: Polish and Integration (2-3 weeks)
1. **React Integration**
   - Table-specific hooks
   - Current row context providers
   - Performance optimization

2. **Testing and Documentation**
   - Comprehensive test suite
   - Excel compatibility testing
   - API documentation and examples

## Breaking Changes and Migration

### API Changes
- No breaking changes to existing cell reference API
- New optional parameters for table-aware operations
- Enhanced error types (additive)

### Parser Changes
- New grammar rules (additive)
- Enhanced AST node types
- Backward compatible with existing formulas

### Performance Considerations
- Table metadata adds minimal memory overhead
- Structured reference resolution is O(1) with proper indexing
- Dependency graph complexity increases linearly with table count

## Testing Strategy

Based on the existing named expression test patterns, table testing should follow a comprehensive multi-layer approach covering all aspects of the table functionality.

### Unit Test Coverage

#### 1. Core Engine Tests (`tests/unit/core/table.test.ts`)
Following the named expression pattern from `engine.test.ts`:

```typescript
describe('Table Management', () => {
  test('should handle table CRUD operations', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Set up header data
    engine.setCellContents({sheet: sheetId, col: 0, row: 0}, 'Name');
    engine.setCellContents({sheet: sheetId, col: 1, row: 0}, 'Sales');
    
    // Test table creation
    const range = {
      start: {sheet: sheetId, col: 0, row: 0}, 
      end: {sheet: sheetId, col: 1, row: 5}
    };
    
    expect(engine.isItPossibleToAddTable('SalesData', range, sheetId)).toBe(true);
    engine.addTable('SalesData', range, sheetId);
    
    // Test table retrieval
    const table = engine.getTable('SalesData', sheetId);
    expect(table?.name).toBe('SalesData');
    expect(table?.range).toEqual(range);
    
    // Test table listing
    const tables = engine.listTables(sheetId);
    expect(tables).toContain('SalesData');
    
    // Test table modification
    const newRange = {
      start: {sheet: sheetId, col: 0, row: 0}, 
      end: {sheet: sheetId, col: 2, row: 8}
    };
    engine.changeTable('SalesData', newRange, sheetId);
    expect(engine.getTable('SalesData', sheetId)?.range).toEqual(newRange);
    
    // Test table removal
    engine.removeTable('SalesData', sheetId);
    expect(engine.listTables(sheetId)).not.toContain('SalesData');
  });

  test('should handle table validation and conflicts', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    const range = {
      start: {sheet: sheetId, col: 0, row: 0}, 
      end: {sheet: sheetId, col: 1, row: 5}
    };
    
    // Test duplicate table names
    engine.addTable('SalesData', range, sheetId);
    expect(engine.isItPossibleToAddTable('SalesData', range, sheetId)).toBe(false);
    
    // Test overwrite option
    const changes = engine.addTable('SalesData', range, sheetId, {overwrite: true});
    expect(changes).toEqual([]);
  });

  test('should get table column names from headers', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Set up table with headers
    engine.setCellContents({sheet: sheetId, col: 0, row: 0}, 'Product');
    engine.setCellContents({sheet: sheetId, col: 1, row: 0}, 'Price');
    engine.setCellContents({sheet: sheetId, col: 2, row: 0}, 'Quantity');
    
    const range = {
      start: {sheet: sheetId, col: 0, row: 0}, 
      end: {sheet: sheetId, col: 2, row: 5}
    };
    engine.addTable('Inventory', range, sheetId);
    
    const columnNames = engine.getTableColumnNames('Inventory', sheetId);
    expect(columnNames).toEqual(['Product', 'Price', 'Quantity']);
  });
});
```

#### 2. Dependency Graph Tests (`tests/unit/evaluator/dependency-graph.test.ts`)
Following the named expression pattern:

```typescript
describe('Table operations', () => {
  test('should add tables to dependency graph', () => {
    const graph = new DependencyGraph();
    
    const tableKey = graph.addTable('SalesData', 0);
    expect(tableKey).toBe('table:0:SalesData');
    
    const tableKey2 = graph.addTable('Inventory', 1);
    expect(tableKey2).toBe('table:1:Inventory');
    
    expect(graph.size).toBe(2);
  });

  test('should create correct table keys', () => {
    expect(DependencyGraph.getTableKey('SalesData', 0)).toBe('table:0:SalesData');
    expect(DependencyGraph.getTableKey('Inventory', 1)).toBe('table:1:Inventory');
  });

  test('should track table dependencies', () => {
    const graph = new DependencyGraph();
    
    const cellKey = graph.addCell({sheet: 0, col: 0, row: 0});
    const tableKey = graph.addTable('SalesData', 0);
    
    graph.addDependency(cellKey, tableKey);
    
    const dependents = graph.getTableDependents(tableKey);
    expect(dependents).toHaveLength(1);
  });
});
```

#### 3. Parser Tests (`tests/unit/parser/structured-reference.test.ts`)
New test file focusing on structured reference parsing:

```typescript
describe('Structured Reference Parsing', () => {
  test('should parse basic table column reference', () => {
    const ast = parseFormula('Table1[Sales]') as StructuredReferenceNode;
    expect(ast.type).toBe('structured-reference');
    expect(ast.tableName).toBe('Table1');
    expect(ast.columnName).toBe('Sales');
    expect(ast.isCurrentRow).toBe(false);
  });

  test('should parse current row reference', () => {
    const ast = parseFormula('[@Sales]') as StructuredReferenceNode;
    expect(ast.type).toBe('structured-reference');
    expect(ast.columnName).toBe('Sales');
    expect(ast.isCurrentRow).toBe(true);
  });

  test('should parse table with selector', () => {
    const ast = parseFormula('Table1[[#Headers],[Sales]]') as StructuredReferenceNode;
    expect(ast.type).toBe('structured-reference');
    expect(ast.tableName).toBe('Table1');
    expect(ast.columnName).toBe('Sales');
    expect(ast.selector).toBe('#Headers');
  });

  test('should distinguish between functions and structured references', () => {
    const func = parseFormula('SUM(Table1[Sales])') as FunctionNode;
    expect(func.type).toBe('function');
    expect(func.children[0].type).toBe('structured-reference');

    const tableRef = parseFormula('Table1[Sales]') as StructuredReferenceNode;
    expect(tableRef.type).toBe('structured-reference');
  });

  test('should extract table references from formulas', () => {
    const tables = extractTableReferences('Table1[Sales]*Table1[Quantity]+Table2[Discount]');
    expect(tables).toEqual(['Table1', 'Table2']);
  });
});
```

#### 4. Evaluator Tests (`tests/unit/evaluator/evaluator.test.ts`)
Adding to existing evaluator tests:

```typescript
describe('Structured references', () => {
  test('should evaluate table column reference', () => {
    // Set up table in context
    context.tables.set('table:0:SalesData', {
      name: 'SalesData',
      range: {start: {sheet: 0, col: 0, row: 0}, end: {sheet: 0, col: 1, row: 3}},
      sheetId: 0
    });
    
    // Mock table data
    context.getRangeValues = jest.fn().mockReturnValue([[100], [200], [300]]);
    
    const node = createTestStructuredReferenceNode('SalesData', 'Amount');
    const result = evaluator.evaluate(node, context);
    
    expect(result.value).toEqual([[100], [200], [300]]);
  });

  test('should handle missing table reference', () => {
    const node = createTestStructuredReferenceNode('UnknownTable', 'Amount');
    const result = evaluator.evaluate(node, context);
    expect(result.value).toBe('#TABLE!');
  });

  test('should handle missing column reference', () => {
    context.tables.set('table:0:SalesData', {
      name: 'SalesData',
      range: {start: {sheet: 0, col: 0, row: 0}, end: {sheet: 0, col: 1, row: 3}},
      sheetId: 0
    });
    
    const node = createTestStructuredReferenceNode('SalesData', 'UnknownColumn');
    const result = evaluator.evaluate(node, context);
    expect(result.value).toBe('#COLUMN!');
  });

  test('should handle current row references', () => {
    context.currentTable = 'SalesData';
    context.currentTableRow = 2;
    context.tables.set('table:0:SalesData', {
      name: 'SalesData',
      range: {start: {sheet: 0, col: 0, row: 0}, end: {sheet: 0, col: 1, row: 3}},
      sheetId: 0
    });
    
    context.getCellValue = jest.fn().mockReturnValue(150);
    
    const node = createTestStructuredReferenceNode('SalesData', 'Amount', true);
    const result = evaluator.evaluate(node, context);
    
    expect(result.value).toBe(150);
  });

  test('should detect current row reference outside table context', () => {
    const node = createTestStructuredReferenceNode('SalesData', 'Amount', true);
    const result = evaluator.evaluate(node, context);
    expect(result.value).toBe('#REF!');
  });
});
```

#### 5. Test Helpers (`tests/unit/evaluator/test-helpers.ts`)
Adding helper functions for table tests:

```typescript
export function createTestStructuredReferenceNode(
  tableName: string,
  columnName?: string,
  isCurrentRow: boolean = false,
  selector?: TableSelector
): StructuredReferenceNode {
  return {
    type: 'structured-reference',
    tableName,
    columnName,
    selector,
    isCurrentRow,
    position: { start: 0, end: 1 }
  };
}

export function createTestTableSelectorNode(
  selector: TableSelector
): TableSelectorNode {
  return {
    type: 'table-selector',
    selector,
    position: { start: 0, end: 1 }
  };
}
```

### Integration Test Coverage

#### 1. Dependency Tracking (`tests/integration/dependency-tracking.test.ts`)
Adding to existing dependency tests:

```typescript
describe('Table dependency tracking', () => {
  test('should track dependencies through table references', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Set up table data
    engine.setCellContents({sheet: sheetId, col: 0, row: 0}, 'Amount');
    engine.setCellContents({sheet: sheetId, col: 0, row: 1}, 100);
    engine.setCellContents({sheet: sheetId, col: 0, row: 2}, 200);
    
    // Create table
    const range = {start: {sheet: sheetId, col: 0, row: 0}, end: {sheet: sheetId, col: 0, row: 2}};
    engine.addTable('SalesData', range, sheetId);
    
    // Create formula referencing table
    engine.setCellContents({sheet: sheetId, col: 1, row: 0}, '=SUM(SalesData[Amount])');
    
    // Test dependency tracking
    const precedents = engine.getCellPrecedents({sheet: sheetId, col: 1, row: 0});
    expect(precedents.length).toBeGreaterThanOrEqual(1);
    
    // Verify table is included in precedents
    const hasTableDep = precedents.some(prec => {
      return 'start' in prec && 'end' in prec; // Range dependency from table
    });
    expect(hasTableDep).toBe(true);
  });

  test('should update dependencies when table is resized', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Set up initial table
    const range = {start: {sheet: sheetId, col: 0, row: 0}, end: {sheet: sheetId, col: 0, row: 2}};
    engine.addTable('SalesData', range, sheetId);
    engine.setCellContents({sheet: sheetId, col: 1, row: 0}, '=SUM(SalesData[Amount])');
    
    // Resize table
    const newRange = {start: {sheet: sheetId, col: 0, row: 0}, end: {sheet: sheetId, col: 0, row: 5}};
    engine.changeTable('SalesData', newRange, sheetId);
    
    // Verify dependencies are updated
    const dependents = engine.getCellDependents({sheet: sheetId, col: 0, row: 4}); // New cell in table
    expect(dependents.length).toBeGreaterThanOrEqual(1);
  });
});
```

#### 2. Table Operations (`tests/integration/table-operations.test.ts`)
Comprehensive integration tests:

```typescript
describe('Table Operations Integration', () => {
  test('should handle complete table workflow', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Set up initial data
    engine.setCellContents({sheet: sheetId, col: 0, row: 0}, [
      ['Product', 'Price', 'Quantity'],
      ['Widget A', 10, 5],
      ['Widget B', 15, 3],
      ['Widget C', 20, 2]
    ]);
    
    // Create table
    const range = {start: {sheet: sheetId, col: 0, row: 0}, end: {sheet: sheetId, col: 2, row: 3}};
    engine.addTable('Inventory', range, sheetId);
    
    // Test structured reference formulas
    engine.setCellContents({sheet: sheetId, col: 3, row: 1}, '=[@Price]*[@Quantity]'); // Current row
    engine.setCellContents({sheet: sheetId, col: 3, row: 4}, '=SUM(Inventory[Price])'); // Column sum
    
    // Verify calculations
    expect(engine.getCellValue({sheet: sheetId, col: 3, row: 1})).toBe(50); // 10*5
    expect(engine.getCellValue({sheet: sheetId, col: 3, row: 4})).toBe(45); // 10+15+20
    
    // Test table resize
    const newRange = {start: {sheet: sheetId, col: 0, row: 0}, end: {sheet: sheetId, col: 2, row: 5}};
    engine.changeTable('Inventory', newRange, sheetId);
    
    // Add new data
    engine.setCellContents({sheet: sheetId, col: 0, row: 4}, 'Widget D');
    engine.setCellContents({sheet: sheetId, col: 1, row: 4}, 25);
    
    // Verify sum updates automatically
    expect(engine.getCellValue({sheet: sheetId, col: 3, row: 4})).toBe(70); // 10+15+20+25
  });

  test('should support all table selectors', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Set up table with headers and data
    engine.setCellContents({sheet: sheetId, col: 0, row: 0}, [
      ['Product', 'Sales'],
      ['A', 100],
      ['B', 200],
      ['C', 300]
    ]);
    
    const range = {start: {sheet: sheetId, col: 0, row: 0}, end: {sheet: sheetId, col: 1, row: 3}};
    engine.addTable('Sales', range, sheetId);
    
    // Test different selectors
    engine.setCellContents({sheet: sheetId, col: 2, row: 0}, '=Sales[[#Headers],[Product]]'); // Header
    engine.setCellContents({sheet: sheetId, col: 2, row: 1}, '=SUM(Sales[[#Data],[Sales]])'); // Data only
    engine.setCellContents({sheet: sheetId, col: 2, row: 2}, '=COUNTA(Sales[[#All],[Product]])'); // All rows
    
    expect(engine.getCellValue({sheet: sheetId, col: 2, row: 0})).toBe('Product');
    expect(engine.getCellValue({sheet: sheetId, col: 2, row: 1})).toBe(600); // 100+200+300
    expect(engine.getCellValue({sheet: sheetId, col: 2, row: 2})).toBe(4); // Header + 3 data
  });
});
```

### Excel Compatibility Tests

#### 1. Excel Comparison (`tests/integration/excel-compatibility/tables.test.ts`)
```typescript
describe('Excel Table Compatibility', () => {
  test('should match Excel structured reference behavior', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Test cases based on Excel behavior
    const excelTestCases = [
      {
        formula: '=SUM(Table1[Sales])',
        expected: 600,
        description: 'Basic column sum'
      },
      {
        formula: '=Table1[@Sales]*1.1',
        expected: 110, // Assuming current row Sales = 100
        description: 'Current row reference with calculation'
      },
      {
        formula: '=Table1[[#Headers],[Sales]]',
        expected: 'Sales',
        description: 'Header reference'
      }
    ];
    
    // Set up test data matching Excel
    // ... setup code ...
    
    excelTestCases.forEach(testCase => {
      const result = engine.calculateFormula(testCase.formula, sheetId);
      expect(result).toBe(testCase.expected);
    });
  });

  test('should handle Excel edge cases', () => {
    // Test error conditions that match Excel
    // Test space handling in table/column names
    // Test case sensitivity behavior
    // Test special character handling
  });
});
```

### Performance Tests

#### 1. Large Table Performance (`tests/integration/performance/table-performance.test.ts`)
```typescript
describe('Table Performance', () => {
  test('should handle large tables efficiently', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    const startTime = performance.now();
    
    // Create large table (1000 rows, 10 columns)
    const largeRange = {
      start: {sheet: sheetId, col: 0, row: 0}, 
      end: {sheet: sheetId, col: 9, row: 999}
    };
    engine.addTable('LargeTable', largeRange, sheetId);
    
    // Add formulas referencing the table
    engine.setCellContents({sheet: sheetId, col: 10, row: 0}, '=SUM(LargeTable[Column1])');
    
    const endTime = performance.now();
    
    // Should complete within reasonable time (adjust threshold as needed)
    expect(endTime - startTime).toBeLessThan(1000); // 1 second
  });
});
```

### Error Handling Tests

#### 1. Error Scenarios (`tests/unit/core/table-errors.test.ts`)
```typescript
describe('Table Error Handling', () => {
  test('should return appropriate error types', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetId = engine.getSheetId(engine.addSheet('Test'));
    
    // Test various error conditions
    expect(() => engine.addTable('', range, sheetId)).toThrow(); // Empty name
    expect(() => engine.addTable('Table1', invalidRange, sheetId)).toThrow(); // Invalid range
    
    // Test formula errors
    const result1 = engine.calculateFormula('=UnknownTable[Column1]', sheetId);
    expect(result1).toBe('#TABLE!');
    
    const result2 = engine.calculateFormula('=Table1[UnknownColumn]', sheetId);
    expect(result2).toBe('#COLUMN!');
  });
});
```

This comprehensive testing strategy ensures tables are thoroughly tested at all levels, following the established patterns from named expressions while covering the unique aspects of table functionality.

## Conclusion

Adding table syntax support to FormulaEngine is a substantial enhancement that requires careful integration across all major systems. The proposed design maintains backward compatibility while adding powerful structured reference capabilities that match Excel's behavior.

Key benefits:
- **Enhanced Readability**: `Table1[Sales]` vs `B2:B100`
- **Automatic Range Updates**: Tables expand/contract automatically
- **Better Maintainability**: Column renames update all references
- **Type Safety**: Structured references are validated at parse time

The phased implementation approach ensures systematic development while maintaining system stability throughout the process.

## Following Named Expression Patterns

After analyzing the current FormulaEngine implementation, tables should follow the exact same patterns as named expressions:

### Storage Pattern
```typescript
// Named expressions use:
private namedExpressions: Map<string, NamedExpression> = new Map();

// Tables should use:
private tables: Map<string, TableDefinition> = new Map();
```

### Key Generation Pattern
```typescript
// Named expressions use:
static getNamedExpressionKey(name: string, scope?: number): string {
  return scope === undefined ? `name:${name}` : `name:${scope}:${name}`;
}

// Tables should use:
static getTableKey(name: string, sheetId: number): string {
  return `table:${sheetId}:${name}`;
}
```

### API Method Pattern
```typescript
// Named expressions have:
isItPossibleToAddNamedExpression(name: string, expression: RawCellContent, scope?: number): boolean
addNamedExpression(name: string, expression: RawCellContent, scope?: number, options?: NamedExpressionOptions): ExportedChange[]
isItPossibleToChangeNamedExpression(name: string, newExpression: RawCellContent, scope?: number): boolean
changeNamedExpression(name: string, newExpression: RawCellContent, scope?: number, options?: NamedExpressionOptions): ExportedChange[]
isItPossibleToRemoveNamedExpression(name: string, scope?: number): boolean
removeNamedExpression(name: string, scope?: number): ExportedChange[]

// Tables should have:
isItPossibleToAddTable(name: string, range: SimpleCellRange, sheetId: number): boolean
addTable(name: string, range: SimpleCellRange, sheetId: number): ExportedChange[]
isItPossibleToChangeTable(name: string, newRange: SimpleCellRange, sheetId: number): boolean
changeTable(name: string, newRange: SimpleCellRange, sheetId: number): ExportedChange[]
isItPossibleToRemoveTable(name: string, sheetId: number): boolean
removeTable(name: string, sheetId: number): ExportedChange[]
```

### Dependency Graph Integration Pattern
```typescript
// Named expressions add themselves to dependency graph:
const namedExprKey = DependencyGraph.getNamedExpressionKey(expressionName, scope);
this.dependencyGraph.addNamedExpression(expressionName, scope);

// Tables should do the same:
const tableKey = DependencyGraph.getTableKey(tableName, sheetId);
this.dependencyGraph.addTable(tableName, sheetId);
```

### Evaluation Context Pattern
```typescript
// Named expressions are in evaluation context:
export interface EvaluationContext {
  namedExpressions: Map<string, NamedExpression>;
  // ...
}

// Tables should be added to evaluation context:
export interface EvaluationContext {
  namedExpressions: Map<string, NamedExpression>;
  tables: Map<string, TableDefinition>;
  // ...
}
```

This ensures tables are treated as first-class citizens in the same way as named expressions, with proper dependency tracking, scoping, and evaluation integration.
