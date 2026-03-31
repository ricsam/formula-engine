/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  type CellAddress,
  type CellStyle,
  type ConditionalStyle,
  type CopyCellsOptions,
  type DirectCellStyle,
  type FiniteSpreadsheetRange,
  type NamedExpression,
  type RangeAddress,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "./types";

import type { FillDirection } from "@ricsam/selection-manager";
import { FormulaEvaluator } from "../evaluator/formula-evaluator";
import { AutoFill } from "./autofill-utils";
import { WorkbookManager } from "./managers/workbook-manager";
import { deserialize, serialize } from "./map-serializer";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameSheetInFormula } from "./sheet-renamer";
import { renameTableInFormula } from "./table-renamer";
import { renameWorkbookInFormula } from "./workbook-renamer";
import {
  cellAddressToKey,
  getCellReference,
  keyToCellAddress,
  parseCellReference,
} from "./utils";
import { CacheManager } from "./managers/cache-manager";
import { NamedExpressionManager } from "./managers/named-expression-manager";
import { TableManager } from "./managers/table-manager";
import { EventManager } from "./managers/event-manager";
import { EvaluationManager } from "./managers/evaluation-manager";
import { DependencyManager } from "./managers/dependency-manager";
import { StyleManager } from "./managers/style-manager";
import { CopyManager } from "./managers/copy-manager";
import { ReferenceManager } from "./managers/reference-manager";
import {
  SchemaManager,
  SchemaValidationError,
} from "./managers/schema-manager";
import type { Schema, CreateSchema, SchemaDeclaration, TableSchemaDefinition, CellSchemaDefinition } from "./schema/schema";
import { buildSchemaFromDeclaration } from "./schema/schema-builder";
import { TableOrm } from "./schema/table-orm";
import { CellOrm } from "./schema/cell-orm";
import { GridOrm } from "./schema/grid-orm";
import type { TableSchemaHeaders } from "./managers/schema-manager";
import {
  ENGINE_SNAPSHOT_VERSION,
  type EngineSnapshotV2,
} from "./engine-snapshot";
import {
  CommandExecutor,
  SchemaIntegrityError,
} from "./commands/command-executor";
import { type EngineAction } from "./commands/types";
import {
  SetCellContentCommand,
  SetSheetContentCommand,
  ClearRangeCommand,
  PasteCellsCommand,
  FillAreasCommand,
  MoveCellCommand,
  MoveRangeCommand,
  AutoFillCommand,
} from "./commands/content-commands";
import {
  AddWorkbookCommand,
  RemoveWorkbookCommand,
  RenameWorkbookCommand,
  CloneWorkbookCommand,
  AddSheetCommand,
  RemoveSheetCommand,
  RenameSheetCommand,
  type StructureCommandDeps,
} from "./commands/structure-commands";
import {
  AddTableCommand,
  RemoveTableCommand,
  RenameTableCommand,
  UpdateTableCommand,
  ResetTablesCommand,
  type TableCommandDeps,
} from "./commands/table-commands";
import {
  AddNamedExpressionCommand,
  RemoveNamedExpressionCommand,
  UpdateNamedExpressionCommand,
  RenameNamedExpressionCommand,
  SetNamedExpressionsCommand,
  type NamedExpressionCommandDeps,
} from "./commands/named-expression-commands";
import {
  SetCellMetadataCommand,
  SetSheetMetadataCommand,
  SetWorkbookMetadataCommand,
} from "./commands/metadata-commands";
import {
  AddConditionalStyleCommand,
  RemoveConditionalStyleCommand,
  AddCellStyleCommand,
  RemoveCellStyleCommand,
  ClearCellStylesCommand,
} from "./commands/style-commands";

type Metadata = {
  cell?: unknown;
  sheet?: unknown;
  workbook?: unknown;
};

type MetadataType<
  TMetadata extends Metadata,
  TKey extends keyof Metadata
> = TMetadata[TKey];

/**
 * Main FormulaEngine class
 * @template TCellMetadata - Consumer-defined type for cell metadata (rich text, links, custom data, etc.)
 * @template TSheetMetadata - Consumer-defined type for sheet metadata (text boxes, frozen panes, etc.)
 * @template TWorkbookMetadata - Consumer-defined type for workbook metadata (themes, document properties, etc.)
 */
export class FormulaEngine<
  TMetadata extends Metadata = Metadata,
  TCreateSchema extends
    | CreateSchema<MetadataType<TMetadata, "cell">, Schema, SchemaDeclaration>
    | undefined = undefined
> {
  private workbookManager: WorkbookManager;
  private namedExpressionManager: NamedExpressionManager;
  private tableManager: TableManager;
  private eventManager: EventManager;
  private evaluationManager: EvaluationManager;
  private autoFillManager: AutoFill;
  private dependencyManager: DependencyManager;
  private styleManager: StyleManager;
  private copyManager: CopyManager;
  private referenceManager: ReferenceManager;
  private schemaManager: SchemaManager;
  private commandExecutor: CommandExecutor;

  public schema: TCreateSchema extends CreateSchema<
    MetadataType<TMetadata, "cell">,
    Schema,
    SchemaDeclaration
  >
    ? TCreateSchema["schema"]
    : undefined;

  private schemaDeclaration: SchemaDeclaration | undefined;

  /**
   * Public access to the store manager for testing
   */
  public _workbookManager: WorkbookManager;
  public _namedExpressionManager: NamedExpressionManager;
  public _tableManager: TableManager;
  public _eventManager: EventManager;
  public _evaluationManager: EvaluationManager;
  public _autoFillManager: AutoFill;
  public _dependencyManager: DependencyManager;
  public _styleManager: StyleManager;

  constructor(schema?: TCreateSchema) {
    this.schemaDeclaration = (schema as any)?.declaration;
    this.eventManager = new EventManager();
    this.workbookManager = new WorkbookManager();
    this.namedExpressionManager = new NamedExpressionManager();
    this.tableManager = new TableManager(this.workbookManager);
    const cacheManager = new CacheManager();
    this.dependencyManager = new DependencyManager(
      cacheManager,
      this.workbookManager
    );

    const formulaEvaluator = new FormulaEvaluator(
      this.tableManager,
      this.dependencyManager,
      this.namedExpressionManager
    );

    this.evaluationManager = new EvaluationManager(
      this.workbookManager,
      this.tableManager,
      formulaEvaluator,
      this.dependencyManager
    );

    this.styleManager = new StyleManager(this.evaluationManager);
    this.copyManager = new CopyManager(
      this.workbookManager,
      this.evaluationManager,
      this.styleManager
    );

    this.autoFillManager = new AutoFill(
      this.workbookManager,
      this.styleManager
    );

    this.referenceManager = new ReferenceManager();
    this.schemaManager = new SchemaManager(this.tableManager);

    // Initialize command executor
    this.commandExecutor = new CommandExecutor(
      this.evaluationManager,
      this.eventManager,
      () =>
        this.schemaManager.validateAllSchemaConstraints(
          (cell) => this.getCellValue(cell),
          (cell) => this.getCellMetadata(cell),
          (table) => this.getTableDataCells(table)
        )
    );

    // Build the working schema from declaration if provided
    if (this.schemaDeclaration) {
      this.schema = buildSchemaFromDeclaration(
        this,
        this.schemaDeclaration,
        this.schemaManager
      ) as any;
    } else {
      this.schema = {} as any;
    }

    this._workbookManager = this.workbookManager;
    this._namedExpressionManager = this.namedExpressionManager;
    this._tableManager = this.tableManager;
    this._eventManager = this.eventManager;
    this._evaluationManager = this.evaluationManager;
    this._autoFillManager = this.autoFillManager;
    this._dependencyManager = this.dependencyManager;
    this._styleManager = this.styleManager;
  }

  /**
   * Static factory method to build an empty engine
   * @template TC - Consumer-defined cell metadata type
   * @template TS - Consumer-defined sheet metadata type
   * @template TW - Consumer-defined workbook metadata type
   */
  static buildEmpty<
    TMetadata extends Metadata = Metadata,
    TSchemaDeclaration extends
      | CreateSchema<MetadataType<TMetadata, "cell">, any, any>
      | undefined = undefined
  >(schema?: TSchemaDeclaration) {
    return new FormulaEngine<TMetadata, TSchemaDeclaration>(schema);
  }

  /**
   * Add a table schema at runtime
   * @param namespace - Unique namespace for the schema
   * @param address - Table address (workbookName and tableName)
   * @param headers - Table headers with parse functions
   * @returns The TableOrm instance for immediate use
   */
  addTableSchema<THeaders extends TableSchemaHeaders<MetadataType<TMetadata, "cell">>>(
    namespace: string,
    address: { workbookName: string; tableName: string },
    headers: THeaders
  ): TableOrm<{
    [K in keyof THeaders]: ReturnType<THeaders[K]["parse"]>;
  }> {
    // Register the schema with the schema manager
    this.schemaManager.registerTableSchema(
      namespace,
      address.workbookName,
      address.tableName,
      headers
    );

    // Create the ORM instance
    const orm = new TableOrm<{
      [K in keyof THeaders]: ReturnType<THeaders[K]["parse"]>;
    }>(
      this,
      address.workbookName,
      address.tableName,
      headers,
      namespace
    );

    // Add to schema object for runtime access
    (this.schema as Record<string, object>)[namespace] = orm;

    return orm;
  }
  

  /**
   * Add a cell schema at runtime
   * @param namespace - Unique namespace for the schema
   * @param cellAddress - Address of the cell
   * @param parse - Parse function for the cell value
   * @returns The CellOrm instance for immediate use
   */
  addCellSchema<TValue>(
    namespace: string,
    cellAddress: CellAddress,
    parse: (value: unknown, metadata: MetadataType<TMetadata, "cell">) => TValue,
    write: (value: TValue) => { value: SerializedCellValue; metadata?: MetadataType<TMetadata, "cell"> } = (value) => ({ value: value as unknown as SerializedCellValue })
  ): CellOrm<TValue> {
    // Register the schema with the schema manager
    this.schemaManager.registerCellSchema(namespace, cellAddress, parse);

    // Create the ORM instance
    const orm = new CellOrm(this, cellAddress, parse, write, namespace);

    // Add to schema object for runtime access
    (this.schema as Record<string, object>)[namespace] = orm;

    return orm;
  }

  /**
   * Add a grid schema at runtime
   * @param namespace - Unique namespace for the schema
   * @param address - Grid address (workbookName and sheetName)
   * @param range - Finite range of cells for the grid
   * @param parse - Parse function for the cell values
   * @param write - Write function for serializing values (optional for primitive types)
   * @returns The GridOrm instance for immediate use
   */
  addGridSchema<TValue>(
    namespace: string,
    address: { workbookName: string; sheetName: string },
    range: FiniteSpreadsheetRange,
    parse: (value: unknown, metadata: MetadataType<TMetadata, "cell">) => TValue,
    write: (value: TValue) => { value: SerializedCellValue; metadata?: MetadataType<TMetadata, "cell"> } = (value) => ({ value: value as unknown as SerializedCellValue })
  ): GridOrm<TValue, MetadataType<TMetadata, "cell">> {
    // Register the schema with the schema manager
    this.schemaManager.registerGridSchema(
      namespace,
      address.workbookName,
      address.sheetName,
      range,
      parse
    );

    // Create the ORM instance
    const orm = new GridOrm<TValue, MetadataType<TMetadata, "cell">>(
      this,
      address.workbookName,
      address.sheetName,
      range,
      parse,
      write,
      namespace
    );

    // Add to schema object for runtime access
    (this.schema as Record<string, object>)[namespace] = orm;

    return orm;
  }

  //#region Cell
  getCellEvaluationResult(
    cellAddress: CellAddress
  ): SingleEvaluationResult | undefined {
    return this.evaluationManager.getCellEvaluationResult(cellAddress);
  }

  getCellValue(cellAddress: CellAddress, debug?: boolean): SerializedCellValue {
    const result = this.getCellEvaluationResult(cellAddress);
    if (!result) {
      return "";
    }

    return this.evaluationManager.evaluationResultToSerializedValue(
      result,
      cellAddress,
      debug
    );
  }

  /**
   * Set metadata for a cell
   * Metadata can contain rich text, links, comments, or any consumer-defined data
   */
  setCellMetadata(
    address: CellAddress,
    metadata: MetadataType<TMetadata, "cell"> | undefined
  ): void {
    this.commandExecutor.execute(
      new SetCellMetadataCommand(this.workbookManager, address, metadata)
    );
  }

  /**
   * Get metadata for a cell
   */
  getCellMetadata(
    address: CellAddress
  ): MetadataType<TMetadata, "cell"> | undefined {
    const metadata = this.workbookManager.getCellMetadata(address);
    return metadata as MetadataType<TMetadata, "cell"> | undefined;
  }

  /**
   * Get all cell metadata for a sheet (serialized as Map)
   */
  getSheetMetadataSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, MetadataType<TMetadata, "sheet">> {
    return this.workbookManager.getSheetMetadataSerialized(opts) as Map<
      string,
      MetadataType<TMetadata, "sheet">
    >;
  }

  /**
   * Set metadata for a sheet
   * Sheet metadata can contain text boxes, frozen panes, print settings, or any consumer-defined data
   */
  setSheetMetadata(
    opts: { workbookName: string; sheetName: string },
    metadata: MetadataType<TMetadata, "sheet">
  ): void {
    this.commandExecutor.execute(
      new SetSheetMetadataCommand(this.workbookManager, opts, metadata)
    );
  }

  /**
   * Get metadata for a sheet
   */
  getSheetMetadata(opts: {
    workbookName: string;
    sheetName: string;
  }): MetadataType<TMetadata, "sheet"> | undefined {
    return this.workbookManager.getSheetMetadata(opts) as
      | MetadataType<TMetadata, "sheet">
      | undefined;
  }

  /**
   * Set metadata for a workbook
   * Workbook metadata can contain themes, document properties, settings, or any consumer-defined data
   */
  setWorkbookMetadata(
    workbookName: string,
    metadata: MetadataType<TMetadata, "workbook">
  ): void {
    this.commandExecutor.execute(
      new SetWorkbookMetadataCommand(
        this.workbookManager,
        workbookName,
        metadata
      )
    );
  }

  /**
   * Get metadata for a workbook
   */
  getWorkbookMetadata(
    workbookName: string
  ): MetadataType<TMetadata, "workbook"> | undefined {
    return this.workbookManager.getWorkbookMetadata(workbookName) as
      | MetadataType<TMetadata, "workbook">
      | undefined;
  }

  //#endregion

  //#region Reference Tracking
  /**
   * Create a tracked reference to a range
   * Returns a stable UUID that can be used to retrieve the address later
   * The reference automatically updates when sheets/workbooks are renamed
   */
  createRef(address: RangeAddress): string {
    return this.referenceManager.createRef(address);
  }

  /**
   * Get the current address for a tracked reference
   * Returns undefined if reference doesn't exist or has been invalidated
   */
  getRefAddress(refId: string): RangeAddress | undefined {
    return this.referenceManager.getRefAddress(refId);
  }

  /**
   * Delete a tracked reference
   * Returns true if the reference was deleted, false if it didn't exist
   */
  deleteRef(refId: string): boolean {
    return this.referenceManager.deleteRef(refId);
  }

  /**
   * Get all invalid reference IDs
   * Useful for cleanup after sheet/workbook deletions
   */
  getInvalidRefs(): string[] {
    return this.referenceManager.getInvalidRefs();
  }
  //#endregion

  evaluateFormula(
    /**
     * formula without the leading = sign
     */
    formula: string,
    cellAddress: CellAddress
  ): SerializedCellValue {
    return this.evaluationManager.evaluateFormula(formula, cellAddress);
  }

  getCellDependents(
    address: CellAddress | SpreadsheetRange
  ): (SpreadsheetRange | CellAddress)[] {
    throw new Error("Not implemented");
  }

  getCellPrecedents(
    address: CellAddress | SpreadsheetRange
  ): (SpreadsheetRange | CellAddress)[] {
    throw new Error("Not implemented");
  }

  //#endregion

  //#region Named Expressions
  addNamedExpression(opts: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }) {
    this.commandExecutor.execute(
      new AddNamedExpressionCommand(this.getNamedExpressionCommandDeps(), opts),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  removeNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    this.commandExecutor.execute(
      new RemoveNamedExpressionCommand(
        this.getNamedExpressionCommandDeps(),
        opts
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Check if a named expression exists
   */
  hasNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): boolean {
    const scope =
      opts.sheetName && opts.workbookName
        ? {
            type: "sheet" as const,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }
        : opts.workbookName
        ? { type: "workbook" as const, workbookName: opts.workbookName }
        : { type: "global" as const };

    return !!this.namedExpressionManager.getNamedExpression({
      name: opts.expressionName,
      scope,
    });
  }

  updateNamedExpression(opts: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    this.commandExecutor.execute(
      new UpdateNamedExpressionCommand(
        this.getNamedExpressionCommandDeps(),
        opts
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  renameNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
    newName: string;
  }): void {
    this.commandExecutor.execute(
      new RenameNamedExpressionCommand(
        this.getNamedExpressionCommandDeps(),
        opts
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  setNamedExpressions(
    opts: (
      | { type: "global" }
      | { type: "sheet"; sheetName: string; workbookName: string }
      | { type: "workbook"; workbookName: string }
    ) & {
      expressions: Map<string, NamedExpression>;
    }
  ) {
    this.commandExecutor.execute(
      new SetNamedExpressionsCommand(
        this.getNamedExpressionCommandDeps(),
        opts
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }
  //#endregion

  //#region Tables
  addTable(props: {
    tableName: string;
    sheetName: string;
    workbookName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }): void {
    this.commandExecutor.execute(
      new AddTableCommand(this.getTableCommandDeps(), props),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  renameTable(
    workbookName: string,
    names: { oldName: string; newName: string }
  ): void {
    this.commandExecutor.execute(
      new RenameTableCommand(
        this.getTableCommandDeps(),
        workbookName,
        names.oldName,
        names.newName
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  updateTable(opts: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
    workbookName: string;
  }): void {
    this.commandExecutor.execute(
      new UpdateTableCommand(this.getTableCommandDeps(), opts),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  removeTable(opts: { tableName: string; workbookName: string }): void {
    this.commandExecutor.execute(
      new RemoveTableCommand(this.getTableCommandDeps(), opts),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Check if a table exists
   */
  hasTable(opts: { tableName: string; workbookName: string }): boolean {
    return !!this.tableManager.getTable({
      workbookName: opts.workbookName,
      name: opts.tableName,
    });
  }

  /**
   * Get a table definition by name
   */
  getTable(opts: {
    tableName: string;
    workbookName: string;
  }): TableDefinition | undefined {
    return this.tableManager.getTable({
      workbookName: opts.workbookName,
      name: opts.tableName,
    });
  }

  resetTables(tables: Map<string, Map<string, TableDefinition>>): void {
    this.commandExecutor.execute(
      new ResetTablesCommand(this.getTableCommandDeps(), tables),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  getTables(workbookName: string) {
    return this.tableManager.getTables(workbookName);
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.tableManager.isCellInTable(cellAddress);
  }

  /**
   * Get all data cells in a table (excluding header row).
   * Since spills cannot enter tables (they get #SPILL! error), this only
   * needs to return cells with direct content.
   */
  private getTableDataCells(table: TableDefinition): CellAddress[] {
    const { start, endRow, headers } = table;
    const dataStartRow = start.rowIndex + 1;
    const endColIndex = start.colIndex + headers.size - 1;

    // Build a RangeAddress for the data area
    const rangeAddress: RangeAddress = {
      workbookName: table.workbookName,
      sheetName: table.sheetName,
      range: {
        start: { col: start.colIndex, row: dataStartRow },
        end: {
          col: { type: "number", value: endColIndex },
          row: endRow,
        },
      },
    };

    // Return cells with direct content - spills cannot enter tables
    return this.workbookManager.getCellsInRange(rangeAddress);
  }

  //#endregion

  //#region Conditional Styling
  /**
   * Add a conditional style rule
   */
  addConditionalStyle(style: ConditionalStyle): void {
    this.commandExecutor.execute(
      new AddConditionalStyleCommand(this.styleManager, style)
    );
  }

  /**
   * Remove a conditional style rule by index
   */
  removeConditionalStyle(workbookName: string, index: number): void {
    this.commandExecutor.execute(
      new RemoveConditionalStyleCommand(this.styleManager, workbookName, index)
    );
  }

  /**
   * Get the count of conditional styles for a workbook
   */
  getConditionalStyleCount(workbookName: string): number {
    const allStyles = this.styleManager.getAllConditionalStyles();
    return allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === workbookName)
    ).length;
  }

  /**
   * Get all conditional styles intersecting with a range
   */
  getConditionalStylesIntersectingWithRange(
    range: RangeAddress
  ): ConditionalStyle[] {
    return this.styleManager.getConditionalStylesIntersectingWithRange(range);
  }

  /**
   * Get the computed style for a specific cell
   */
  getCellStyle(cellAddress: CellAddress): CellStyle | undefined {
    return this.styleManager.getCellStyle(cellAddress);
  }

  /**
   * Get all cell styles (for testing and serialization)
   */
  getAllCellStyles(): DirectCellStyle[] {
    return this.styleManager.getAllCellStyles();
  }

  /**
   * Get all conditional styles (for testing and serialization)
   */
  getAllConditionalStyles(): ConditionalStyle[] {
    return this.styleManager.getAllConditionalStyles();
  }

  /**
   * Add a direct cell style rule
   */
  addCellStyle(style: DirectCellStyle): void {
    this.commandExecutor.execute(
      new AddCellStyleCommand(this.styleManager, style)
    );
  }

  /**
   * Remove a direct cell style rule by index
   */
  removeCellStyle(workbookName: string, index: number): void {
    this.commandExecutor.execute(
      new RemoveCellStyleCommand(this.styleManager, workbookName, index)
    );
  }

  /**
   * Get the count of direct cell styles for a workbook
   */
  getCellStyleCount(workbookName: string): number {
    const allStyles = this.styleManager.getAllCellStyles();
    return allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === workbookName)
    ).length;
  }

  /**
   * Get all direct cell styles intersecting with a range
   */
  getStylesIntersectingWithRange(range: RangeAddress): DirectCellStyle[] {
    return this.styleManager.getStylesIntersectingWithRange(range);
  }

  /**
   * Get the style for a range if all cells in the range have the same style
   * Returns the DirectCellStyle if the range is completely contained within a single style's area
   * Returns undefined if multiple styles, partial coverage, or no styles apply
   */
  getStyleForRange(range: RangeAddress): DirectCellStyle | undefined {
    return this.styleManager.getStyleForRange(range);
  }

  /**
   * Clear all cell styles and conditional styles for a given range
   * Adjusts existing style ranges rather than deleting them entirely
   */
  clearCellStyles(range: RangeAddress): void {
    this.commandExecutor.execute(
      new ClearCellStylesCommand(this.styleManager, range)
    );
  }

  //#endregion

  //#region Copy/Paste
  /**
   * Paste cells from source to target
   */
  pasteCells(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): void {
    this.commandExecutor.execute(
      new PasteCellsCommand(
        this.workbookManager,
        this.copyManager,
        source,
        target,
        options
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Fill one or more areas with a seed range's content/style
   * Uses column-first strategy: fills down, then replicates right
   * Formulas are adjusted based on each target cell's offset from the seed
   *
   * @param seedRange - The range to use as a template/pattern
   * @param targetRanges - One or more range addresses to fill
   * @param options - Copy options (target: 'all'|'content'|'style', type: 'value'|'formula', cut: boolean)
   *
   * @example
   * // Fill F6:J10 with A1:B2 seed (2x2 pattern fills 5x5 area)
   * engine.fillAreas(
   *   {
   *     workbookName,
   *     sheetName,
   *     range: {
   *       start: { col: 0, row: 0 },
   *       end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
   *     }
   *   },
   *   [{
   *     workbookName,
   *     sheetName,
   *     range: {
   *       start: { col: 5, row: 5 },
   *       end: { col: { type: "number", value: 9 }, row: { type: "number", value: 9 } }
   *     }
   *   }],
   *   { cut: false, type: "formula", target: "all" }
   * );
   */
  fillAreas(
    seedRange: RangeAddress,
    targetRanges: RangeAddress[],
    options: CopyCellsOptions
  ): void {
    this.commandExecutor.execute(
      new FillAreasCommand(
        this.workbookManager,
        this.copyManager,
        seedRange,
        targetRanges,
        options
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Smart paste that automatically determines whether to paste or fill
   * Handles multiple selection areas - each area is independently pasted or filled
   * - If area is larger than source, uses fillAreas() to fill the area
   * - If area is same size or smaller, uses pasteCells() for normal paste
   *
   * @param sourceCells - The copied cells
   * @param pasteSelection - One or more selection areas where user is pasting
   * @param options - Copy options
   *
   * @example
   * // Copy A1, paste into two areas B1:C2 and E5:F6 - both get filled
   * engine.smartPaste(
   *   [{ workbookName, sheetName, colIndex: 0, rowIndex: 0 }],
   *   {
   *     workbookName,
   *     sheetName,
   *     areas: [
   *       { start: { col: 1, row: 0 }, end: { col: { type: "number", value: 2 }, row: { type: "number", value: 1 } } },
   *       { start: { col: 4, row: 4 }, end: { col: { type: "number", value: 5 }, row: { type: "number", value: 5 } } }
   *     ]
   *   },
   *   { cut: false, type: "formula", target: "all" }
   * );
   */
  smartPaste(
    sourceCells: CellAddress[],
    pasteSelection: {
      workbookName: string;
      sheetName: string;
      areas: SpreadsheetRange[];
    },
    options: CopyCellsOptions
  ): void {
    if (sourceCells.length === 0) {
      return;
    }

    // If cut operation, always use pasteCells (never fillAreas)
    // Cut should be a simple move operation, not a fill
    if (options.cut === true) {
      for (const area of pasteSelection.areas) {
        const target: CellAddress = {
          workbookName: pasteSelection.workbookName,
          sheetName: pasteSelection.sheetName,
          colIndex: area.start.col,
          rowIndex: area.start.row,
        };
        this.pasteCells(sourceCells, target, options);
      }
      return;
    }

    // For copy operations (not cut), use smart paste/fill logic
    // Calculate source bounds once
    const sourceBounds = this.getBoundsFromCells(sourceCells);
    const sourceWidth = sourceBounds.maxCol - sourceBounds.minCol + 1;
    const sourceHeight = sourceBounds.maxRow - sourceBounds.minRow + 1;

    // Create seed range for fill operations
    const seedRange: RangeAddress = {
      workbookName: sourceCells[0]!.workbookName,
      sheetName: sourceCells[0]!.sheetName,
      range: {
        start: { col: sourceBounds.minCol, row: sourceBounds.minRow },
        end: {
          col: { type: "number", value: sourceBounds.maxCol },
          row: { type: "number", value: sourceBounds.maxRow },
        },
      },
    };

    // Process each selected area independently
    for (const area of pasteSelection.areas) {
      const pasteStartCol = area.start.col;
      const pasteStartRow = area.start.row;
      const pasteEndCol =
        area.end.col.type === "number" ? area.end.col.value : pasteStartCol;
      const pasteEndRow =
        area.end.row.type === "number" ? area.end.row.value : pasteStartRow;

      const pasteWidth = pasteEndCol - pasteStartCol + 1;
      const pasteHeight = pasteEndRow - pasteStartRow + 1;

      // Decide per area: paste or fill?
      const shouldFill = pasteWidth > sourceWidth || pasteHeight > sourceHeight;

      if (shouldFill) {
        // Use fillAreas for this area
        const targetRange: RangeAddress = {
          workbookName: pasteSelection.workbookName,
          sheetName: pasteSelection.sheetName,
          range: {
            start: { col: pasteStartCol, row: pasteStartRow },
            end: {
              col: { type: "number", value: pasteEndCol },
              row: { type: "number", value: pasteEndRow },
            },
          },
        };

        this.fillAreas(seedRange, [targetRange], options);
      } else {
        // Use pasteCells for this area
        const target: CellAddress = {
          workbookName: pasteSelection.workbookName,
          sheetName: pasteSelection.sheetName,
          colIndex: pasteStartCol,
          rowIndex: pasteStartRow,
        };

        this.pasteCells(sourceCells, target, options);
      }
    }
  }

  /**
   * Get bounds (min/max row/col) from an array of cell addresses
   */
  private getBoundsFromCells(cells: CellAddress[]): {
    minCol: number;
    minRow: number;
    maxCol: number;
    maxRow: number;
  } {
    if (cells.length === 0) {
      throw new Error("Cannot get bounds from empty cell array");
    }

    let minCol = Infinity;
    let minRow = Infinity;
    let maxCol = -Infinity;
    let maxRow = -Infinity;

    for (const cell of cells) {
      minCol = Math.min(minCol, cell.colIndex);
      minRow = Math.min(minRow, cell.rowIndex);
      maxCol = Math.max(maxCol, cell.colIndex);
      maxRow = Math.max(maxRow, cell.rowIndex);
    }

    return { minCol, minRow, maxCol, maxRow };
  }

  /**
   * Move a single cell to a new location
   * Updates all formula references that point to the moved cell
   *
   * @param source - The cell to move
   * @param target - The destination cell address
   *
   * @example
   * // Move A1 to D5. If B1 contains =A1, it will be updated to =D5
   * engine.moveCell(
   *   { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
   *   { workbookName, sheetName, colIndex: 3, rowIndex: 4 }
   * );
   */
  moveCell(source: CellAddress, target: CellAddress): void {
    this.commandExecutor.execute(
      new MoveCellCommand(
        this.workbookManager,
        this.copyManager,
        source,
        target
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Move a range of cells to a new location
   * Updates all formula references that point to the moved cells
   *
   * @param sourceRange - The range to move
   * @param target - The top-left destination cell address
   *
   * @example
   * // Move A1:D5 to F10. If E1 contains =SUM(A1:D5), it will be updated to =SUM(F10:I14)
   * engine.moveRange(
   *   {
   *     workbookName,
   *     sheetName,
   *     range: {
   *       start: { col: 0, row: 0 },
   *       end: { col: { type: "number", value: 3 }, row: { type: "number", value: 4 } }
   *     }
   *   },
   *   { workbookName, sheetName, colIndex: 5, rowIndex: 9 }
   * );
   */
  moveRange(sourceRange: RangeAddress, target: CellAddress): void {
    this.commandExecutor.execute(
      new MoveRangeCommand(
        this.workbookManager,
        this.copyManager,
        sourceRange,
        target
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }
  //#endregion

  //#region Sheets
  addSheet(opts: { workbookName: string; sheetName: string }): void {
    this.commandExecutor.execute(
      new AddSheetCommand(this.getStructureCommandDeps(), opts),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  removeSheet(opts: { workbookName: string; sheetName: string }): void {
    this.commandExecutor.execute(
      new RemoveSheetCommand(this.getStructureCommandDeps(), opts),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  renameSheet(opts: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    this.commandExecutor.execute(
      new RenameSheetCommand(this.getStructureCommandDeps(), opts),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Check if a sheet exists
   */
  hasSheet(opts: { workbookName: string; sheetName: string }): boolean {
    return !!this.workbookManager.getSheet(opts);
  }

  getSheets(workbookName: string) {
    return this.workbookManager.getSheets(workbookName);
  }

  getSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }) {
    return this.workbookManager.getSheet({ workbookName, sheetName });
  }

  getSheetSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, SerializedCellValue> {
    return this.workbookManager.getSheetSerialized(opts);
  }

  //#endregion

  //#region Workbook
  addWorkbook(workbookName: string): void {
    this.commandExecutor.execute(
      new AddWorkbookCommand(this.getStructureCommandDeps(), workbookName),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  removeWorkbook(workbookName: string): void {
    this.commandExecutor.execute(
      new RemoveWorkbookCommand(this.getStructureCommandDeps(), workbookName),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Check if a workbook exists
   */
  hasWorkbook(workbookName: string): boolean {
    return this.workbookManager.getWorkbooks().has(workbookName);
  }

  cloneWorkbook(fromWorkbookName: string, toWorkbookName: string): void {
    this.commandExecutor.execute(
      new CloneWorkbookCommand(
        this.getStructureCommandDeps(),
        fromWorkbookName,
        toWorkbookName
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  renameWorkbook(opts: { workbookName: string; newWorkbookName: string }) {
    this.commandExecutor.execute(
      new RenameWorkbookCommand(
        this.getStructureCommandDeps(),
        opts.workbookName,
        opts.newWorkbookName
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  getWorkbooks() {
    return this.workbookManager.getWorkbooks();
  }
  //#endregion

  //#region CRUD Operations
  /**
   * Overrides the content of a sheet.
   * @param sheetName - The name of the sheet to set the content of
   * @param content - A map of cell addresses to their serialized values
   * @remarks This method is used to set the content of a sheet. It will re-evaluate all sheets to ensure all dependencies are resolved correctly.
   * @throws SchemaIntegrityError if any evaluated cell value violates a schema constraint
   */
  setSheetContent(
    opts: { sheetName: string; workbookName: string },
    content: Map<string, SerializedCellValue>
  ) {
    this.commandExecutor.execute(
      new SetSheetContentCommand(this.workbookManager, opts, content),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Set the content of a single cell.
   * @throws SchemaIntegrityError if the evaluated value violates a schema constraint
   */
  setCellContent(address: CellAddress, content: SerializedCellValue) {
    this.commandExecutor.execute(
      new SetCellContentCommand(this.workbookManager, address, content),
      { validate: this.schemaManager.hasSchemas() }
    );
  }
  //#endregion

  //#region Auto-fill
  /**
   * Auto-fills one or more ranges based on the seedRange and the direction.
   * Supports pattern detection and style copying.
   */
  autoFill(
    opts: { sheetName: string; workbookName: string },
    /**
     * The user's original selection that defines the pattern/series.
     */
    seedRange: SpreadsheetRange,
    /**
     * One or more ranges to fill (the new cells populated by the drag, excluding the seed)
     */
    fillRanges: SpreadsheetRange[],
    /**
     * The direction of the fill.
     */
    direction: FillDirection
  ): void {
    this.commandExecutor.execute(
      new AutoFillCommand(
        this.workbookManager,
        this.styleManager,
        this.autoFillManager,
        opts,
        seedRange,
        fillRanges,
        direction
      ),
      { validate: this.schemaManager.hasSchemas() }
    );
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   */
  clearSpreadsheetRange(address: RangeAddress) {
    this.commandExecutor.execute(
      new ClearRangeCommand(this.workbookManager, address),
      { validate: this.schemaManager.hasSchemas() }
    );
  }
  //#endregion

  //#region State - UI library integration
  getState() {
    return {
      workbooks: this.workbookManager.getWorkbooks(),
      namedExpressions: this.namedExpressionManager.getNamedExpressions(),
      tables: this.tableManager.tables,
      conditionalStyles: this.styleManager.getAllConditionalStyles(),
      cellStyles: this.styleManager.getAllCellStyles(),
      references: this.referenceManager.getAllReferences(),
    };
  }

  onUpdate(listener: () => void) {
    return this.eventManager.onUpdate(listener);
  }

  private buildSerializedSnapshot(): EngineSnapshotV2 {
    const evaluationSnapshots = this.dependencyManager.toSnapshot(
      this.evaluationManager
    );

    return {
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        workbook: this.workbookManager.toSnapshot(),
        namedExpression: this.namedExpressionManager.toSnapshot(),
        table: this.tableManager.toSnapshot(),
        style: this.styleManager.toSnapshot(),
        reference: this.referenceManager.toSnapshot(),
        dependency: evaluationSnapshots.dependency,
        cache: evaluationSnapshots.cache,
      },
    };
  }

  serializeEngine(): string {
    return serialize(this.buildSerializedSnapshot());
  }

  resetToSerializedEngine(data: string) {
    const deserialized = deserialize(data) as Partial<EngineSnapshotV2>;
    if (
      !deserialized ||
      typeof deserialized !== "object" ||
      deserialized.version !== ENGINE_SNAPSHOT_VERSION ||
      !deserialized.managers
    ) {
      throw new Error(
        "Unsupported serialized engine format. Expected EngineSnapshot version 2."
      );
    }

    this.workbookManager.restoreFromSnapshot(deserialized.managers.workbook);

    deserialized.managers.workbook.forEach((workbook) => {
      this.namedExpressionManager.addWorkbook(workbook.name);
      workbook.sheets.forEach((sheet) => {
        this.namedExpressionManager.addSheet({
          workbookName: workbook.name,
          sheetName: sheet.name,
        });
      });
    });

    this.namedExpressionManager.restoreFromSnapshot(
      deserialized.managers.namedExpression
    );
    this.tableManager.restoreFromSnapshot(deserialized.managers.table);
    this.styleManager.restoreFromSnapshot(deserialized.managers.style);
    this.referenceManager.restoreFromSnapshot(deserialized.managers.reference);
    this.dependencyManager.restoreFromSnapshot(
      {
        dependency: deserialized.managers.dependency,
        cache: deserialized.managers.cache,
      },
      this.evaluationManager
    );
    this.commandExecutor.clearHistory();
    this.commandExecutor.clearActionLog();

    this.eventManager.emitUpdate();
  }
  //#endregion

  //#region Undo/Redo
  /**
   * Undo the last command.
   * @returns true if undo was performed, false if nothing to undo
   */
  undo(): boolean {
    return this.commandExecutor.undo();
  }

  /**
   * Redo the last undone command.
   * @returns true if redo was performed, false if nothing to redo
   */
  redo(): boolean {
    return this.commandExecutor.redo();
  }

  /**
   * Check if undo is available.
   */
  canUndo(): boolean {
    return this.commandExecutor.canUndo();
  }

  /**
   * Check if redo is available.
   */
  canRedo(): boolean {
    return this.commandExecutor.canRedo();
  }

  /**
   * Get the action log for persistence/collaboration.
   * Actions can be serialized and replayed to reconstruct state.
   */
  getActionLog(): EngineAction[] {
    return this.commandExecutor.getActionLog();
  }

  /**
   * Clear the action log.
   */
  clearActionLog(): void {
    this.commandExecutor.clearActionLog();
  }

  /**
   * Clear undo/redo history.
   */
  clearHistory(): void {
    this.commandExecutor.clearHistory();
  }
  //#endregion

  //#region Command Dependencies (internal)
  /**
   * Get dependencies for structure commands.
   * @internal
   */
  private getStructureCommandDeps(): StructureCommandDeps {
    return {
      workbookManager: this.workbookManager,
      namedExpressionManager: this.namedExpressionManager,
      tableManager: this.tableManager,
      styleManager: this.styleManager,
      referenceManager: this.referenceManager,
      apiSchemaManager: this.schemaManager,
      renameSheetInFormula,
      renameWorkbookInFormula,
    };
  }

  /**
   * Get dependencies for table commands.
   * @internal
   */
  private getTableCommandDeps(): TableCommandDeps {
    return {
      tableManager: this.tableManager,
      namedExpressionManager: this.namedExpressionManager,
      workbookManager: this.workbookManager,
      apiSchemaManager: this.schemaManager,
      getCellValue: (cell) => this.getCellValue(cell),
      renameTableInFormula,
    };
  }

  /**
   * Get dependencies for named expression commands.
   * @internal
   */
  private getNamedExpressionCommandDeps(): NamedExpressionCommandDeps {
    return {
      namedExpressionManager: this.namedExpressionManager,
      workbookManager: this.workbookManager,
      renameNamedExpressionInFormula,
    };
  }
  //#endregion
}
