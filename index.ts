/**
 * FormulaEngine - A TypeScript-based spreadsheet formula evaluation library
 */

// Export the main engine
export { FormulaEngine } from './src/core/engine';

// Export core types
export type {
  // Cell addressing
  SimpleCellAddress,
  SimpleCellRange,
  
  // Cell values and types
  CellValue,
  RawCellContent,
  CellType,
  CellValueType,
  CellValueDetailedType,
  FormulaError,
  
  // Sheet structure
  BoundingRect,
  FormatInfo,
  
  // Changes and events
  ExportedChange,
  FormulaEngineEvents,
  
  // Named expressions
  NamedExpression,
  SerializedNamedExpression,
  NamedExpressionOptions,
  
  // Configuration
  FormulaEngineOptions,
  
  // Results
  Result
} from './src/core/types';

// Export utility functions
export {
  // Address conversion
  colNumberToLetter,
  letterToColNumber,
  addressToKey,
  keyToAddress,
  
  // Type guards
  isFormulaError,
  isCellEmpty,
  isNumber,
  isString,
  isBoolean,
  getCellValueType
} from './src/core/types';

// Export address utilities
export {
  parseCellAddress,
  parseCellRange,
  addressToA1,
  rangeToA1,
  isAddressInRange,
  doRangesOverlap,
  getRangeSize,
  offsetAddress,
  offsetRange,
  expandRangeToInclude,
  iterateRange,
  adjustReferences,
  isValidAddress,
  createRange
} from './src/core/address';

// Version
export const VERSION = '0.1.0';