/**
 * Command Pattern Types for FormulaEngine
 *
 * Commands encapsulate all mutating operations on the engine,
 * enabling undo/redo, schema validation with rollback, and action serialization.
 */

import type { CellAddress, RangeAddress, SerializedCellValue } from "../types";

/**
 * Serializable action representation of a command.
 * Used for persistence, collaboration, and changelog functionality.
 */
export interface EngineAction {
  type: string;
  payload: unknown;
  timestamp?: number;
}

/**
 * Base interface for all engine commands.
 */
export interface EngineCommand {
  /**
   * Whether this command affects cell values/formulas and requires re-evaluation.
   * Commands that only affect metadata or styles don't need re-evaluation.
   */
  readonly requiresReevaluation: boolean;

  /**
   * Execute the command (forward operation).
   */
  execute(): void;

  /**
   * Undo the command (reverse operation).
   */
  undo(): void;

  /**
   * Convert the command to a serializable action for persistence/collaboration.
   */
  toAction(): EngineAction;
}

/**
 * Options for command execution.
 */
export interface ExecuteOptions {
  /**
   * Whether to validate schema constraints after execution.
   * Only applies to commands that require re-evaluation.
   */
  validate?: boolean;

  /**
   * Whether to skip adding to undo stack (for internal use).
   */
  skipUndoStack?: boolean;

  /**
   * Whether to skip emitting update events.
   */
  skipEmitUpdate?: boolean;
}

/**
 * Result of schema validation.
 */
export interface SchemaValidationResult {
  valid: boolean;
  errors: SchemaValidationErrorInfo[];
}

/**
 * Information about a schema validation error.
 */
export interface SchemaValidationErrorInfo {
  message: string;
  cellAddress?: CellAddress;
  schemaNamespace?: string;
  columnName?: string;
  originalError?: Error;
}

/**
 * Action types for all commands.
 * Used for serialization and deserialization.
 */
export const ActionTypes = {
  // Content commands
  SET_CELL_CONTENT: "SET_CELL_CONTENT",
  SET_SHEET_CONTENT: "SET_SHEET_CONTENT",
  CLEAR_RANGE: "CLEAR_RANGE",
  PASTE_CELLS: "PASTE_CELLS",
  FILL_AREAS: "FILL_AREAS",
  MOVE_CELL: "MOVE_CELL",
  MOVE_RANGE: "MOVE_RANGE",
  AUTO_FILL: "AUTO_FILL",

  // Structure commands - Workbook
  ADD_WORKBOOK: "ADD_WORKBOOK",
  REMOVE_WORKBOOK: "REMOVE_WORKBOOK",
  RENAME_WORKBOOK: "RENAME_WORKBOOK",
  CLONE_WORKBOOK: "CLONE_WORKBOOK",

  // Structure commands - Sheet
  ADD_SHEET: "ADD_SHEET",
  REMOVE_SHEET: "REMOVE_SHEET",
  RENAME_SHEET: "RENAME_SHEET",

  // Table commands
  ADD_TABLE: "ADD_TABLE",
  REMOVE_TABLE: "REMOVE_TABLE",
  RENAME_TABLE: "RENAME_TABLE",
  UPDATE_TABLE: "UPDATE_TABLE",
  RESET_TABLES: "RESET_TABLES",

  // Named expression commands
  ADD_NAMED_EXPRESSION: "ADD_NAMED_EXPRESSION",
  REMOVE_NAMED_EXPRESSION: "REMOVE_NAMED_EXPRESSION",
  UPDATE_NAMED_EXPRESSION: "UPDATE_NAMED_EXPRESSION",
  RENAME_NAMED_EXPRESSION: "RENAME_NAMED_EXPRESSION",
  SET_NAMED_EXPRESSIONS: "SET_NAMED_EXPRESSIONS",

  // Metadata commands
  SET_CELL_METADATA: "SET_CELL_METADATA",
  SET_SHEET_METADATA: "SET_SHEET_METADATA",
  SET_WORKBOOK_METADATA: "SET_WORKBOOK_METADATA",

  // Style commands
  ADD_CONDITIONAL_STYLE: "ADD_CONDITIONAL_STYLE",
  REMOVE_CONDITIONAL_STYLE: "REMOVE_CONDITIONAL_STYLE",
  ADD_CELL_STYLE: "ADD_CELL_STYLE",
  REMOVE_CELL_STYLE: "REMOVE_CELL_STYLE",
  CLEAR_CELL_STYLES: "CLEAR_CELL_STYLES",

  // State commands
  RESET_TO_SERIALIZED: "RESET_TO_SERIALIZED",
} as const;

export type ActionType = (typeof ActionTypes)[keyof typeof ActionTypes];

