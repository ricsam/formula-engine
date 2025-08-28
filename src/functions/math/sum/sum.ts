import {
  FormulaError,
  type CellAddress,
  type CellInfinity,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";

/**
 * SUM function - Adds all numbers in the arguments
 */
export const SUM: FunctionDefinition = {
  name: "SUM",
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
          message: `Can't sum non-number (${result.result.type}, ${result.result.value})`,
        };
      }
      const range = result.spillArea;
      let subTotal = 0;
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
              if (parsed.type === "error" || parsed.type === "infinity") {
                return parsed;
              }
              subTotal += parsed.value;
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
              if (
                parsedSpillResult.type === "error" ||
                parsedSpillResult.type === "infinity"
              ) {
                return parsedSpillResult;
              }
              subTotal += parsedSpillResult.value;
            }
          }
        }
      }

      if (
        range.end.col.type === "infinity" &&
        range.end.row.type === "infinity"
      ) {
        throw new Error("SUM over an infinite end-range is not implemented. TODO");
      }

      if (range.end.col.type === "infinity") {
        throw new Error("SUM over an infinite col-range is not implemented. TODO");
      }

      if (range.end.row.type === "infinity") {
        throw new Error("SUM over an infinite row-range is not implemented. TODO");
      }

      return {
        type: "number",
        value: subTotal,
      };
    };
    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      let total = 0;
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);
        const parsed = parseResult(value);
        if (parsed.type === "error" || parsed.type === "infinity") {
          return parsed;
        }
        total += parsed.value;
      }
      return {
        type: "number",
        value: total,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
