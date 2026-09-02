/**
 * FormulaEngine - A TypeScript-based spreadsheet formula evaluation library
 */

// Export the main engine
export { FormulaEngine } from "./src/core/engine";
export type {
  CellAddress,
  FormulaEngineOptions,
  RangeAddress,
  ReplaceChange,
  ReplaceTarget,
  SearchMatch,
  SearchOptions,
  SpreadsheetRange,
  UndoRedoOptions,
  UndoRedoState,
} from "./src/core/types";

export {
  analyzeFormula,
  findFormulaReferenceAt,
} from "./src/language/formula-analysis";
export type {
  FormulaAnalysis,
  FormulaAnalysisOptions,
  FormulaDiagnostic,
  FormulaReference,
  FormulaReferenceResolution,
  FormulaReferenceTarget,
  FormulaToken,
  FormulaTokenKind,
  FormulaTokenModifier,
  NamedExpressionScope,
  SourceSpan,
} from "./src/language/formula-analysis";

// Export React integration (optional - only import if using React)
export { useEngine as useSerializedSheet } from "./src/react/hooks";

// Version
export const VERSION = "0.1.0";
