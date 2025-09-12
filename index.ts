/**
 * FormulaEngine - A TypeScript-based spreadsheet formula evaluation library
 */

// Export the main engine
export { FormulaEngine } from "./src/core/engine";

// Export React integration (optional - only import if using React)
export { useEngine as useSerializedSheet } from "./src/react/hooks";

// Version
export const VERSION = "0.1.0";
