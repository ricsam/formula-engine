import {
  FormulaError,
  type CellAddress,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SpreadsheetRange,
} from "src/core/types";

/**
 * SEQUENCE(rows, [columns], [start], [step])
 * Generates a sequence of numbers in an array.
 * 
 * @param rows - The number of rows to return (can be array - uses origin value)
 * @param columns - [Optional] The number of columns to return (default: 1, can be array - uses origin value)
 * @param start - [Optional] The first number in the sequence (default: 1, can be array - uses origin value)
 * @param step - [Optional] The amount to increment each subsequent value (default: 1, can be array - uses origin value)
 * 
 * Returns a spilled array of sequential numbers.
 * When any argument is an array, the result is broadcast over the largest input spill area.
 */
export const SEQUENCE: FunctionDefinition = {
  name: "SEQUENCE",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 1 || node.args.length > 4) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "SEQUENCE function takes 1 to 4 arguments",
      };
    }

    // Evaluate all arguments and check for array inputs
    const argResults: FunctionEvaluationResult[] = [];
    let largestSpillArea: SpreadsheetRange | null = null;
    let hasArrayInput = false;

    for (let i = 0; i < node.args.length; i++) {
      const result = this.evaluateNode(node.args[i]!, context);
      if (result.type === "error") {
        return result;
      }
      argResults.push(result);

      if (result.type === "spilled-values") {
        hasArrayInput = true;
        if (!largestSpillArea) {
          largestSpillArea = result.spillArea;
        } else {
          // Use the larger spill area (more cells)
          const currentCols = result.spillArea.end.col.type === "number" ? result.spillArea.end.col.value : Infinity;
          const currentRows = result.spillArea.end.row.type === "number" ? result.spillArea.end.row.value : Infinity;
          const currentSize = (currentCols - result.spillArea.start.col + 1) *
                             (currentRows - result.spillArea.start.row + 1);
          
          const largestCols = largestSpillArea.end.col.type === "number" ? largestSpillArea.end.col.value : Infinity;
          const largestRows = largestSpillArea.end.row.type === "number" ? largestSpillArea.end.row.value : Infinity;
          const largestSize = (largestCols - largestSpillArea.start.col + 1) *
                             (largestRows - largestSpillArea.start.row + 1);
          
          if (currentSize > largestSize) {
            largestSpillArea = result.spillArea;
          }
        }
      }
    }

    // Extract values from arguments (using origin values for arrays)
    const rowsResult = argResults[0];
    if (!rowsResult) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Missing rows argument",
      };
    }
    
    let rowsValue: number;
    
    if (rowsResult.type === "spilled-values") {
      if (rowsResult.originResult.type !== "number") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Rows argument must be a number",
        };
      }
      rowsValue = rowsResult.originResult.value;
    } else if (rowsResult.type === "value" && rowsResult.result.type === "number") {
      rowsValue = rowsResult.result.value;
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Rows argument must be a number",
      };
    }
    
    const rows = Math.floor(rowsValue);
    if (rows < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Rows must be greater than 0",
      };
    }

    // Evaluate columns argument (optional, default: 1)
    let columns = 1;
    if (node.args.length > 1) {
      const columnsResult = argResults[1];
      if (!columnsResult) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Missing columns argument",
        };
      }
      
      if (columnsResult.type === "spilled-values") {
        if (columnsResult.originResult.type !== "number") {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Columns argument must be a number",
          };
        }
        columns = Math.floor(columnsResult.originResult.value);
      } else if (columnsResult.type === "value" && columnsResult.result.type === "number") {
        columns = Math.floor(columnsResult.result.value);
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Columns argument must be a number",
        };
      }
      if (columns < 1) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Columns must be greater than 0",
        };
      }
    }

    // Evaluate start argument (optional, default: 1)
    let start = 1;
    if (node.args.length > 2) {
      const startResult = argResults[2];
      if (!startResult) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Missing start argument",
        };
      }
      
      if (startResult.type === "spilled-values") {
        if (startResult.originResult.type !== "number") {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Start argument must be a number",
          };
        }
        start = startResult.originResult.value;
      } else if (startResult.type === "value" && startResult.result.type === "number") {
        start = startResult.result.value;
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Start argument must be a number",
        };
      }
    }

    // Evaluate step argument (optional, default: 1)
    let step = 1;
    if (node.args.length > 3) {
      const stepResult = argResults[3];
      if (!stepResult) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Missing step argument",
        };
      }
      
      if (stepResult.type === "spilled-values") {
        if (stepResult.originResult.type !== "number") {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Step argument must be a number",
          };
        }
        step = stepResult.originResult.value;
      } else if (stepResult.type === "value" && stepResult.result.type === "number") {
        step = stepResult.result.value;
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Step argument must be a number",
        };
      }
    }

    // Create the spill area - if there are array inputs, use the largest input spill area
    // Otherwise, create a new spill area based on the SEQUENCE dimensions
    const spillArea: SpreadsheetRange = hasArrayInput && largestSpillArea ? largestSpillArea : {
      start: {
        col: context.currentCell.colIndex,
        row: context.currentCell.rowIndex,
      },
      end: {
        col: { type: "number", value: context.currentCell.colIndex + columns - 1 },
        row: { type: "number", value: context.currentCell.rowIndex + rows - 1 },
      },
    };

    // Calculate the origin result (top-left cell value)
    const originResult: CellNumber = {
      type: "number",
      value: start,
    };

    return {
      type: "spilled-values",
      spillArea,
      spillOrigin: context.currentCell,
      source: "SEQUENCE function",
      originResult,
      evaluate: (spilledCell, context) => {
        if (hasArrayInput) {
          // When we have array inputs, all cells in the broadcast area get the same SEQUENCE result
          // The SEQUENCE is calculated once using the origin values from the input arrays
          
          // For broadcasting, we always return the same sequence across all spilled cells
          // Calculate the value for the origin cell of the SEQUENCE (always (0,0))
          const x = 0;
          const y = 0;
          
          // Check if the origin cell is within our intended SEQUENCE area
          if (x >= columns || y >= rows) {
            return undefined;
          }
          
          // Calculate the sequential value for origin cell
          const sequenceIndex = y * columns + x;
          const value = start + (sequenceIndex * step);
          
          return {
            type: "value",
            result: {
              type: "number",
              value,
            },
          };
        } else {
          // Normal SEQUENCE behavior - calculate the value for this spilled cell based on its offset
          const x = spilledCell.spillOffset.x;
          const y = spilledCell.spillOffset.y;
          
          // Check if the spilled cell is within our intended area
          if (x >= columns || y >= rows) {
            return undefined;
          }
          
          // Calculate the sequential value
          // Values are filled row by row: (0,0), (0,1), (0,2), ..., (1,0), (1,1), etc.
          const sequenceIndex = y * columns + x;
          const value = start + (sequenceIndex * step);
          
          return {
            type: "value",
            result: {
              type: "number",
              value,
            },
          };
        }
      },
    };
  },
};