/**
 * Commands Module - Command Pattern implementation for FormulaEngine
 *
 * This module provides:
 * - EngineCommand interface for all commands
 * - CommandExecutor for execution with undo/redo and schema validation
 * - All command implementations organized by category
 */

// Types and executor
export * from "./types";
export { CommandExecutor, SchemaIntegrityError } from "./command-executor";

// Content commands
export {
  SetCellContentCommand,
  SetSheetContentCommand,
  ClearRangeCommand,
  PasteCellsCommand,
  FillAreasCommand,
  MoveCellCommand,
  MoveRangeCommand,
  AutoFillCommand,
} from "./content-commands";

// Structure commands
export {
  AddWorkbookCommand,
  RemoveWorkbookCommand,
  RenameWorkbookCommand,
  CloneWorkbookCommand,
  AddSheetCommand,
  RemoveSheetCommand,
  RenameSheetCommand,
  type StructureCommandDeps,
} from "./structure-commands";

// Table commands
export {
  AddTableCommand,
  RemoveTableCommand,
  RenameTableCommand,
  UpdateTableCommand,
  ResetTablesCommand,
  type TableCommandDeps,
} from "./table-commands";

// Named expression commands
export {
  AddNamedExpressionCommand,
  RemoveNamedExpressionCommand,
  UpdateNamedExpressionCommand,
  RenameNamedExpressionCommand,
  SetNamedExpressionsCommand,
  type NamedExpressionCommandDeps,
} from "./named-expression-commands";

// Metadata commands
export {
  SetCellMetadataCommand,
  SetSheetMetadataCommand,
  SetWorkbookMetadataCommand,
} from "./metadata-commands";

// Style commands
export {
  AddConditionalStyleCommand,
  RemoveConditionalStyleCommand,
  AddCellStyleCommand,
  RemoveCellStyleCommand,
  ClearCellStylesCommand,
} from "./style-commands";

