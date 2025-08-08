/**
 * FormulaEngine - A TypeScript-based spreadsheet formula evaluation library
 */

// Export the main engine
export { FormulaEngine } from "./src/core/engine";

// Export core types
export type {
  // Sheet structure
  BoundingRect,
  CellType,
  // Cell values and types
  CellValue,
  CellValueDetailedType,
  CellValueType,
  // Changes and events
  ExportedChange,
  FormatInfo,
  FormulaEngineEvents,
  // Configuration
  FormulaEngineOptions,
  FormulaError,
  // Named expressions
  NamedExpression,
  NamedExpressionOptions,
  RawCellContent,
  // Results
  Result,
  SerializedNamedExpression,
  // Cell addressing
  SimpleCellAddress,
  SimpleCellRange,
} from "./src/core/types";

// Export utility functions
export {
  addressToKey,
  // Address conversion
  colNumberToLetter,
  getCellValueType,
  isBoolean,
  isCellEmpty,
  // Type guards
  isFormulaError,
  isNumber,
  isString,
  keyToAddress,
  letterToColNumber,
} from "./src/core/types";

// Export address utilities
export {
  addressToA1,
  adjustReferences,
  createRange,
  doRangesOverlap,
  expandRangeToInclude,
  getRangeSize,
  isAddressInRange,
  isValidAddress,
  iterateRange,
  offsetAddress,
  offsetRange,
  parseCellAddress,
  parseCellRange,
  rangeToA1,
} from "./src/core/address";

// Export React integration (optional - only import if using React)
export { useSerializedSheet } from "./src/react/hooks";

// Version
export const VERSION = "0.1.0";
