import {
  FormulaError,
  type CellAddress,
  type CellInfinity,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";

/**
 * MAX function - Returns the largest number in the arguments
 */
export const MAX: FunctionDefinition = {
  name: "MAX",
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
          message: `Can't find max of non-number (${result.result.type}, ${result.result.value})`,
        };
      }
      const range = result.spillArea;
      let maxValue = -Infinity;
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
                if (parsed.sign === "positive") {
                  return parsed; // Positive infinity is always the maximum
                }
                // Negative infinity - continue checking other values
              } else {
                maxValue = Math.max(maxValue, parsed.value);
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
                if (parsedSpillResult.sign === "positive") {
                  return parsedSpillResult; // Positive infinity is always the maximum
                }
                // Negative infinity - continue checking other values
              } else {
                maxValue = Math.max(maxValue, parsedSpillResult.value);
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
          message: "Cannot find maximum of empty range",
        };
      }

      return {
        type: "number",
        value: maxValue,
      };
    };

    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      let maxValue = -Infinity;
      let hasValues = false;
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);
        const parsed = parseResult(value);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          if (parsed.sign === "positive") {
            return parsed; // Positive infinity is always the maximum
          }
          // Negative infinity - continue checking other values
        } else {
          maxValue = Math.max(maxValue, parsed.value);
          hasValues = true;
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "MAX requires at least one numeric argument",
        };
      }

      return {
        type: "number",
        value: maxValue,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
