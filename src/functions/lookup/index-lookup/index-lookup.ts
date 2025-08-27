import {
  FormulaError,
  type CellValue,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type SpilledValuesEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * INDEX function - Returns a value from a table or array
 * INDEX(array, row_num, [column_num])
 * 
 * STRICT TYPE CHECKING:
 * - array: range/array only
 * - row_num: number only (integer, 1-based)
 * - column_num: number only (integer, 1-based, optional)
 * 
 * If array is 1-dimensional (single row or column):
 * - Only row_num is used (treats as a linear array)
 * - column_num is ignored if provided
 * 
 * If array is 2-dimensional:
 * - Both row_num and column_num can be used
 * - If column_num is omitted, returns entire row
 */

// Helper function to get array dimensions from a spilled result
function getArrayDimensions(spillArea: SpreadsheetRange): {
  rows: number;
  cols: number;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const startRow = spillArea.start.row;
  const startCol = spillArea.start.col;
  const endRow = spillArea.end.row.type === "number" ? spillArea.end.row.value : startRow;
  const endCol = spillArea.end.col.type === "number" ? spillArea.end.col.value : startCol;
  
  return {
    rows: endRow - startRow + 1,
    cols: endCol - startCol + 1,
    startRow,
    startCol,
    endRow,
    endCol,
  };
}

// Helper function to get value from array at specific position
function getValueFromArray(
  this: FormulaEvaluator,
  arrayResult: SpilledValuesEvaluationResult,
  row: number,
  col: number,
  context: EvaluationContext
): CellValue | { type: "error"; err: FormulaError; message: string } {
  const dims = getArrayDimensions(arrayResult.spillArea);
  
  // Convert 1-based indices to 0-based
  const rowIndex = row - 1;
  const colIndex = col - 1;
  
  // Check bounds
  if (rowIndex < 0 || rowIndex >= dims.rows) {
    return {
      type: "error",
      err: FormulaError.REF,
      message: `INDEX: row_num ${row} is out of range (1-${dims.rows})`,
    };
  }
  
  if (colIndex < 0 || colIndex >= dims.cols) {
    return {
      type: "error",
      err: FormulaError.REF,
      message: `INDEX: column_num ${col} is out of range (1-${dims.cols})`,
    };
  }
  
  // Calculate actual cell position
  const actualRow = dims.startRow + rowIndex;
  const actualCol = dims.startCol + colIndex;
  
  const spilledAddress: CellAddress = {
    colIndex: actualCol,
    rowIndex: actualRow,
    sheetName: context.currentSheet,
  };
  
  const spill = {
    address: spilledAddress,
    spillOffset: {
      x: actualCol - dims.startCol,
      y: actualRow - dims.startRow,
    },
  };
  
  const spillResult = arrayResult.evaluate(spill, context);
  
  if (!spillResult || spillResult.type !== "value") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "INDEX: Unable to retrieve value from array",
    };
  }
  
  return spillResult.result;
}

export const INDEX: FunctionDefinition = {
  name: "INDEX",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDEX function takes 2 or 3 arguments",
      };
    }

    // Evaluate array argument
    const arrayResult = this.evaluateNode(node.args[0]!, context);
    if (arrayResult.type === "error") {
      return arrayResult;
    }

    // Evaluate row_num argument
    const rowNumResult = this.evaluateNode(node.args[1]!, context);
    if (rowNumResult.type === "error") {
      return rowNumResult;
    }

    // Evaluate column_num argument (optional)
    let colNumResult: FunctionEvaluationResult | null = null;
    if (node.args[2]) {
      colNumResult = this.evaluateNode(node.args[2], context);
      if (colNumResult.type === "error") {
        return colNumResult;
      }
    }

    // Handle spilled arrays for row_num and col_num (not array which is expected to be a range)
    if (
      rowNumResult.type === "spilled-values" ||
      (colNumResult && colNumResult.type === "spilled-values")
    ) {
      // TODO: Implement comprehensive spilled array support like FIND function
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDEX: Spilled row_num/column_num arguments not yet implemented",
      };
    }

    // Validate argument types
    if (rowNumResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDEX: Invalid row_num result type",
      };
    }

    if (colNumResult && colNumResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDEX: Invalid column_num result type",
      };
    }

    // Strict type checking for row_num
    if (rowNumResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `INDEX row_num must be number, got ${rowNumResult.result.type}`,
      };
    }

    // Strict type checking for column_num if provided
    if (colNumResult && colNumResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `INDEX column_num must be number, got ${colNumResult.result.type}`,
      };
    }

    // Extract row and column numbers (convert to integers)
    const rowNum = Math.floor((rowNumResult.result as { type: "number"; value: number }).value);
    const colNum = colNumResult ? Math.floor((colNumResult.result as { type: "number"; value: number }).value) : 1;

    // Validate that indices are positive
    if (rowNum < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `INDEX row_num must be >= 1, got ${rowNum}`,
      };
    }

    if (colNum < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `INDEX column_num must be >= 1, got ${colNum}`,
      };
    }

    // Handle different array result types
    if (arrayResult.type === "value") {
      // Single value case - can only access [1,1]
      if (rowNum !== 1 || colNum !== 1) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `INDEX: Single value can only be accessed at [1,1], got [${rowNum},${colNum}]`,
        };
      }
      return {
        type: "value",
        result: arrayResult.result,
      } satisfies ValueEvaluationResult;
    } else if (arrayResult.type === "spilled-values") {
      // Array case - use helper function to get value
      const result = getValueFromArray.call(this, arrayResult, rowNum, colNum, context);
      
      if (result.type === "error") {
        return result;
      }
      
      return {
        type: "value",
        result,
      } satisfies ValueEvaluationResult;
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INDEX: Invalid array argument type",
      };
    }
  },
};
