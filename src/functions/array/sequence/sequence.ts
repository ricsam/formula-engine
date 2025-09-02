import {
  FormulaError,
  type CellAddress,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SpreadsheetRange,
  type PositiveInfinity,
  type ErrorEvaluationResult,
} from "src/core/types";
import { getCellReference } from "src/core/utils";
import {
  getRangeIntersection,
  OpenRangeEvaluator,
} from "src/functions/math/open-range-evaluator";

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
        const spillArea = result.spillArea(context.currentCell);
        if (!largestSpillArea) {
          largestSpillArea = spillArea;
        } else {
          // Use the larger spill area (more cells)
          const currentCols =
            spillArea.end.col.type === "number"
              ? spillArea.end.col.value
              : Infinity;
          const currentRows =
            spillArea.end.row.type === "number"
              ? spillArea.end.row.value
              : Infinity;
          const currentSize =
            (currentCols - spillArea.start.col + 1) *
            (currentRows - spillArea.start.row + 1);

          const largestCols =
            largestSpillArea.end.col.type === "number"
              ? largestSpillArea.end.col.value
              : Infinity;
          const largestRows =
            largestSpillArea.end.row.type === "number"
              ? largestSpillArea.end.row.value
              : Infinity;
          const largestSize =
            (largestCols - largestSpillArea.start.col + 1) *
            (largestRows - largestSpillArea.start.row + 1);

          if (currentSize > largestSize) {
            largestSpillArea = spillArea;
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

    let rowsValue: number | "infinity";
    let isRowsInfinite = false;

    if (rowsResult.type === "spilled-values") {
      throw new Error("Sequences cannot contain spilled values");
    } else if (rowsResult.type === "value") {
      if (rowsResult.result.type === "infinity") {
        rowsValue = "infinity";
        isRowsInfinite = true;
      } else if (rowsResult.result.type === "number") {
        rowsValue = rowsResult.result.value;
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Rows argument must be a number or INFINITY",
        };
      }
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Rows argument must be a number or INFINITY",
      };
    }

    const rows = isRowsInfinite ? Infinity : Math.floor(rowsValue as number);
    if (!isRowsInfinite && rows < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Rows must be greater than 0",
      };
    }

    // Evaluate columns argument (optional, default: 1)
    let columns = 1;
    let isColumnsInfinite = false;
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
        throw new Error("Sequences cannot contain spilled values");
      } else if (columnsResult.type === "value") {
        if (columnsResult.result.type === "infinity") {
          columns = Infinity;
          isColumnsInfinite = true;
        } else if (columnsResult.result.type === "number") {
          columns = Math.floor(columnsResult.result.value);
        } else {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Columns argument must be a number or INFINITY",
          };
        }
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Columns argument must be a number or INFINITY",
        };
      }
      if (!isColumnsInfinite && columns < 1) {
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
        throw new Error("Sequences cannot contain spilled values");
      } else if (
        startResult.type === "value" &&
        startResult.result.type === "number"
      ) {
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
        throw new Error("Sequences cannot contain spilled values");
      } else if (
        stepResult.type === "value" &&
        stepResult.result.type === "number"
      ) {
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
    const spillArea = (origin: CellAddress): SpreadsheetRange => {
      return hasArrayInput && largestSpillArea
        ? largestSpillArea
        : {
            start: {
              col: origin.colIndex,
              row: origin.rowIndex,
            },
            end: {
              col: isColumnsInfinite
                ? { type: "infinity", sign: "positive" }
                : { type: "number", value: origin.colIndex + columns - 1 },
              row: isRowsInfinite
                ? { type: "infinity", sign: "positive" }
                : { type: "number", value: origin.rowIndex + rows - 1 },
            },
          };
    };

    // Calculate the origin result (top-left cell value)
    const originResult: CellNumber = {
      type: "number",
      value: start,
    };

    return {
      type: "spilled-values",
      spillArea,
      source: "SEQUENCE function",
      evaluate: (spillOffset, context) => {
        if (hasArrayInput) {
          // When we have array inputs, all cells in the broadcast area get the same SEQUENCE result
          // The SEQUENCE is calculated once using the origin values from the input arrays

          // For broadcasting, we always return the same sequence across all spilled cells
          // Calculate the value for the origin cell of the SEQUENCE (always (0,0))
          const x = 0;
          const y = 0;

          // Check if the origin cell is within our intended SEQUENCE area
          if (
            (!isColumnsInfinite && x >= columns) ||
            (!isRowsInfinite && y >= rows)
          ) {
            return undefined;
          }

          // Calculate the sequential value for origin cell
          // Since this is for broadcasting, we always use the origin (0,0) value
          const sequenceIndex = 0;
          const value = start + sequenceIndex * step;

          return {
            type: "value",
            result: {
              type: "number",
              value,
            },
          };
        } else {
          // Normal SEQUENCE behavior - calculate the value for this spilled cell based on its offset
          const x = spillOffset.x;
          const y = spillOffset.y;

          // Check if the spilled cell is within our intended area
          if (
            (!isColumnsInfinite && x >= columns) ||
            (!isRowsInfinite && y >= rows)
          ) {
            return undefined;
          }

          // Calculate the sequential value
          // Values are filled row by row: (0,0), (0,1), (0,2), ..., (1,0), (1,1), etc.
          let sequenceIndex: number;
          if (isColumnsInfinite && isRowsInfinite) {
            // For infinite 2D sequences, we can't calculate a single index
            // We'll use a simple pattern based on position
            sequenceIndex = y + x;
          } else if (isColumnsInfinite) {
            // For infinite columns, sequence goes horizontally
            sequenceIndex = x;
          } else if (isRowsInfinite) {
            // For infinite rows, sequence goes vertically
            sequenceIndex = y;
          } else {
            // Finite dimensions - normal calculation
            sequenceIndex = y * columns + x;
          }

          const value = start + sequenceIndex * step;

          return {
            type: "value",
            result: {
              type: "number",
              value,
            },
          };
        }
      },
      evaluateAllCells: function* ({ evaluate, intersection, context, origin }) {
        let range = spillArea(origin);
        if (intersection) {
          const newRange = getRangeIntersection(range, intersection);
          if (!newRange) {
            yield {
              type: "error",
              err: FormulaError.REF,
              message: "Intersection is not valid #2",
            };
            return;
          }
          range = newRange;
        }
        if (
          range.end.row.type === "infinity" ||
          range.end.col.type === "infinity"
        ) {
          const hasIntersection = intersection !== undefined;
          yield {
            type: "error",
            err: FormulaError.REF,
            message: `Can not evaluate all cells over an infinite range`,
          };
          return;
        }

        for (let i = range.start.row; i <= range.end.row.value; i++) {
          for (let j = range.start.col; j <= range.end.col.value; j++) {
            const offsetLeft = j - origin.colIndex;
            const offsetTop = i - origin.rowIndex;

            const evaled = evaluate({ x: offsetLeft, y: offsetTop }, context);
            yield evaled ?? {
              type: "error",
              err: FormulaError.REF,
              message: "Error evaluating SEQUENCE",
            };
          }
        }
      },
    };
  },
};
