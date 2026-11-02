import {
  type CellAddress,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type ValueEvaluationResult,
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { ASTNode } from "../../../parser/ast";
import type { ParsedCriteria } from "../../criteria-parser";

/**
 * Performs counting of numeric values from an iterable of SingleEvaluationResult values
 * Handles numbers and infinities according to Excel's COUNT behavior
 * - Propagates errors immediately
 * - Only counts numeric values (numbers and infinities)
 * - Ignores non-numeric values (strings, booleans, empty cells)
 *
 * @param results - Iterable of SingleEvaluationResult to count
 * @returns FunctionEvaluationResult with the count
 */
export function performCount(
  results: Iterable<SingleEvaluationResult>
): FunctionEvaluationResult {
  let count = 0;

  for (const result of results) {
    if (result.type === "error" || result.type === "awaiting-evaluation") {
      // Propagate errors immediately
      return result;
    }

    if (result.type === "value") {
      if (
        result.result.type === "number" ||
        result.result.type === "infinity"
      ) {
        count++;
      }
      // Non-numeric values (strings, booleans) are ignored
    }
  }

  return {
    type: "value",
    result: { type: "number", value: count },
  };
}

/**
 * Whether a directly supplied argument is a reference to an empty cell.
 *
 * References carry the address they read (`sourceCell`), and the engine reads an
 * unpopulated cell as empty text. A referenced formula that returns "" reads the
 * same, so it counts as empty too; inside a range the two are told apart.
 */
function isEmptyCellReference(result: ValueEvaluationResult): boolean {
  return (
    result.sourceCell !== undefined &&
    result.result.type === "string" &&
    result.result.value === ""
  );
}

/**
 * Counts the values that are not empty in COUNTA's arguments.
 *
 * - A directly supplied value counts once, whatever it is: a number, text
 *   (including the literal ""), a logical value or an error value. A reference
 *   to an empty cell counts nothing.
 * - A spilled argument (a range, or a result such as SEQUENCE or UNIQUE) counts
 *   the cells its `evaluateAllCells` returns. For a raw range those are only the
 *   populated cells plus values spilled onto the range, so a cell whose formula
 *   returns "" counts and an unpopulated cell does not. Errors inside the array
 *   are counted, not propagated; open-ended ranges such as `A:A` work because
 *   only their populated cells are read.
 * - An argument whose cells cannot be read (evaluateAllCells returns an error)
 *   makes the count unknown, so that error is returned, as are
 *   awaiting-evaluation results.
 */
export function countNonEmptyValues(
  evaluator: FormulaEvaluator,
  args: readonly ASTNode[],
  context: EvaluationContext
): FunctionEvaluationResult {
  let count = 0;

  for (const arg of args) {
    const result = evaluator.evaluateNode(arg, context);

    switch (result.type) {
      case "awaiting-evaluation":
        return result;
      case "error":
        count++;
        break;
      case "value":
        if (!isEmptyCellReference(result)) {
          count++;
        }
        break;
      case "spilled-values": {
        const cells = result.evaluateAllCells.call(evaluator, {
          context,
          evaluate: result.evaluate,
          origin: context.cellAddress,
          lookupOrder: "col-major",
        });
        if (cells.type !== "values") {
          return cells;
        }
        for (const cell of cells.values) {
          if (cell.result.type === "awaiting-evaluation") {
            return cell.result;
          }
          count++;
        }
        break;
      }
    }
  }

  return {
    type: "value",
    result: { type: "number", value: count },
  };
}

export function countEmptyCells(
  criteriaRangeResult: FunctionEvaluationResult,
  parsedCriteria: ParsedCriteria,
  currentCell: CellAddress
): FunctionEvaluationResult | undefined {
  // Special case: counting empty cells over infinite ranges
  if (
    criteriaRangeResult.type === "spilled-values" &&
    parsedCriteria.type === "exact" &&
    parsedCriteria.value.type === "string" &&
    parsedCriteria.value.value === ""
  ) {
    const spillArea = criteriaRangeResult.spillArea(currentCell);

    // Check if this is an infinite range
    if (
      spillArea.end.col.type === "infinity" ||
      spillArea.end.row.type === "infinity"
    ) {
      // Return infinity for infinite empty cell count
      return {
        type: "value",
        result: { type: "infinity", sign: "positive" },
      };
    }
  }
}
