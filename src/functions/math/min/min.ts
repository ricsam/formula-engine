import {
  FormulaError,
  type CellAddress,
  type CellInfinity,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";

/**
 * MIN function - Returns the smallest number in the arguments
 */
export const MIN: FunctionDefinition = {
  name: "MIN",
  evaluate: function (node, context) {
    const parseResult = (
      result: FunctionEvaluationResult
    ):
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      if (result.type === "error") {
        return result;
      }
      if (result.type === "value") {
        if (
          result.result.type === "number" ||
          result.result.type === "infinity"
        ) {
          return result.result;
        }
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `Can't find min of non-number (${result.result.type}, ${result.result.value})`,
        };
      }
      const range = result.spillArea;
      let minValue = Infinity;
      let hasValues = false;
      if (range.end.col.type === "number" && range.end.row.type === "number") {
        for (let row = range.start.row; row <= range.end.row.value; row++) {
          for (let col = range.start.col; col <= range.end.col.value; col++) {
            if (
              row === result.spillOrigin.rowIndex &&
              col === result.spillOrigin.colIndex
            ) {
              const parsed = parseResult({
                type: "value",
                result: result.originResult,
              });
              if (parsed.type === "error") {
                return parsed;
              }
              if (parsed.type === "infinity") {
                if (parsed.sign === "negative") {
                  return parsed; // Negative infinity is always the minimum
                }
                // Positive infinity - continue checking other values
              } else {
                minValue = Math.min(minValue, parsed.value);
                hasValues = true;
              }
              continue;
            }
            const spilledAddress: CellAddress = {
              colIndex: col,
              rowIndex: row,
              sheetName: result.spillOrigin.sheetName,
            };
            const spill = {
              address: spilledAddress,
              spillOffset: {
                x: col - result.spillOrigin.colIndex,
                y: row - result.spillOrigin.rowIndex,
              },
            };
            const spillResult = result.evaluate(spill, context);

            if (spillResult) {
              const parsedSpillResult = parseResult(spillResult);
              if (parsedSpillResult.type === "error") {
                return parsedSpillResult;
              }
              if (parsedSpillResult.type === "infinity") {
                if (parsedSpillResult.sign === "negative") {
                  return parsedSpillResult; // Negative infinity is always the minimum
                }
                // Positive infinity - continue checking other values
              } else {
                minValue = Math.min(minValue, parsedSpillResult.value);
                hasValues = true;
              }
            }
          }
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Cannot find minimum of empty range",
        };
      }

      return {
        type: "number",
        value: minValue,
      };
    };

    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      let minValue = Infinity;
      let hasValues = false;
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);
        const parsed = parseResult(value);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          if (parsed.sign === "negative") {
            return parsed; // Negative infinity is always the minimum
          }
          // Positive infinity - continue checking other values
        } else {
          minValue = Math.min(minValue, parsed.value);
          hasValues = true;
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "MIN requires at least one numeric argument",
        };
      }

      return {
        type: "number",
        value: minValue,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
