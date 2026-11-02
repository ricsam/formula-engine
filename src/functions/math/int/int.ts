import {
  FormulaError,
  type CellValue,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpilledValuesEvaluationResult,
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";

/**
 * Rounds one value down to the nearest integer, or explains why it cannot.
 */
function roundDown(
  value: CellValue,
  context: EvaluationContext
): SingleEvaluationResult {
  switch (value.type) {
    case "number": {
      const rounded = Math.floor(value.value);
      // Math.floor(-0) is -0, which spreadsheets do not have.
      return {
        type: "value",
        result: { type: "number", value: rounded === 0 ? 0 : rounded },
      };
    }
    case "infinity":
      // Rounding an infinity down leaves it unchanged, as CEILING does.
      return { type: "value", result: value };
    default:
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `INT number must be a number, got ${
          value.type === "string" && value.value === ""
            ? "an empty value"
            : value.type
        }`,
        errAddress: context.dependencyNode,
      };
  }
}

/**
 * Maps INT over every cell of a spilled argument.
 */
function roundDownSpilled(
  numbers: SpilledValuesEvaluationResult
): SpilledValuesEvaluationResult {
  return {
    type: "spilled-values",
    spillArea: (origin) => numbers.spillArea(origin),
    source: "INT function",
    evaluate: (spillOffset, context) => {
      const cell = numbers.evaluate(spillOffset, context);
      if (cell.type !== "value") {
        return cell;
      }
      return roundDown(cell.result, context);
    },
    evaluateAllCells: function (options) {
      // The argument's cells are produced by the argument's own evaluate; ours
      // would round them before this mapping rounds them again.
      const cells = numbers.evaluateAllCells.call(this, {
        ...options,
        evaluate: numbers.evaluate,
      });
      if (cells.type !== "values") {
        return cells;
      }
      return {
        type: "values",
        values: cells.values.map((cell) => ({
          relativePos: cell.relativePos,
          result:
            cell.result.type === "value"
              ? roundDown(cell.result.result, options.context)
              : cell.result,
        })),
      };
    },
  };
}

/**
 * INT(number)
 *
 * Rounds a number down to the nearest integer, toward negative infinity:
 * `INT(8.9)` is 8 and `INT(-8.9)` is -9.
 *
 * - Arguments follow the coercion rules of this engine's other math functions
 *   (CEILING and the arithmetic operators): only numbers are accepted. Text,
 *   logical values and empty cells, which the engine reads as empty text, give
 *   #VALUE!. This deviates from Excel, which coerces `INT(TRUE)` to 1,
 *   `INT("3.7")` to 3 and an empty cell to 0.
 * - INFINITY and -INFINITY are returned unchanged, matching `CEILING(INFINITY, 1)`.
 * - Errors propagate.
 * - An array argument (a range, a spilled result such as `SEQUENCE(4)/2`, or a
 *   reference to a spilling cell) is rounded element by element and spills with
 *   the argument's shape; an error in one element only affects that element.
 *   Aggregates read the rounded cells through `evaluateAllCells`, which rounds
 *   the cells the argument itself provides: for a range, its populated cells.
 */
export const INT: FunctionDefinition = {
  name: "INT",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "INT function takes exactly 1 argument",
        errAddress: context.dependencyNode,
      };
    }

    const numberResult = this.evaluateNode(node.args[0]!, context);
    if (
      numberResult.type === "error" ||
      numberResult.type === "awaiting-evaluation"
    ) {
      return numberResult;
    }

    if (numberResult.type === "spilled-values") {
      return roundDownSpilled(numberResult);
    }

    return roundDown(numberResult.result, context);
  },
};
