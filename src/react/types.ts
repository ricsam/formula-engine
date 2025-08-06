/**
 * React integration types for FormulaEngine
 */

import type { CellValue, SimpleCellAddress } from '../core/types';

export interface UseSpreadsheetOptions {
  /**
   * Enable automatic re-rendering when the sheet changes
   * @default true
   */
  autoUpdate?: boolean;
  
  /**
   * Debounce delay for updates in milliseconds
   * @default 0
   */
  debounceMs?: number;
}

export interface UseCellOptions {
  /**
   * Enable automatic re-rendering when the cell changes
   * @default true
   */
  autoUpdate?: boolean;
  
  /**
   * Debounce delay for updates in milliseconds
   * @default 0
   */
  debounceMs?: number;
}

export interface UseSpreadsheetRangeOptions {
  /**
   * Enable automatic re-rendering when cells in the range change
   * @default true
   */
  autoUpdate?: boolean;
  
  /**
   * Debounce delay for updates in milliseconds
   * @default 0
   */
  debounceMs?: number;
}

export interface SpreadsheetHookResult {
  /**
   * Map of cell addresses to values for the entire sheet
   */
  spreadsheet: Map<string, CellValue>;
  
  /**
   * Indicates if the hook is loading initial data
   */
  isLoading: boolean;
  
  /**
   * Error state if any
   */
  error: Error | null;
}

export interface CellHookResult {
  /**
   * The current value of the cell
   */
  value: CellValue;
  
  /**
   * Indicates if the hook is loading initial data
   */
  isLoading: boolean;
  
  /**
   * Error state if any
   */
  error: Error | null;
}

export interface SpreadsheetRangeHookResult {
  /**
   * Map of cell addresses to values for the specified range
   */
  rangeData: Map<string, CellValue>;
  
  /**
   * Indicates if the hook is loading initial data
   */
  isLoading: boolean;
  
  /**
   * Error state if any
   */
  error: Error | null;
}
