import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../../../core/types";
import {
  createBroadcastResult,
  toBroadcastOperand,
} from "../../../evaluator/broadcast";
import { getSuppliedArgument } from "../../../parser/ast";
import { isTruthy } from "../if/if";

/**
 * IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)
 *
 * Returns the value that belongs to the first true condition, like nested IFs:
 * `IFS(A1>=90, "A", A1>=80, "B", TRUE, "C")` is
 * `IF(A1>=90, "A", IF(A1>=80, "B", "C"))`.
 *
 * - Conditions are tested like IF's: FALSE, 0 and empty text (an empty cell)
 *   are false, every other value is true. Excel gives #VALUE! for a text
 *   condition; this engine treats non-empty text as true, as IF does.
 * - When no condition is true the result is #N/A. Add `TRUE, fallback` as the
 *   last pair for a default.
 * - Single-value conditions are evaluated in order, and only up to the first
 *   true one; only its value is evaluated, and it is returned as is, so a range
 *   or array value spills. An error in a condition that is reached is returned.
 * - An array condition makes IFS choose cell by cell, like IF with an array
 *   condition: from that pair on, the conditions and values are broadcast
 *   against each other (see `evaluator/broadcast.ts`). Each cell reads its
 *   conditions in order and only the value it chooses, so an error in a
 *   single-value condition or value shows up only in the cells that reach it.
 *   A cell without a true condition is #N/A.
 * - The arguments come in pairs: an odd number of arguments, or none, gives
 *   #VALUE! (Excel rejects such a formula). An omitted condition is FALSE and an
 *   omitted value is 0, as in IF.
 */
export const IFS: FunctionDefinition = {
  name: "IFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    const argCount = node.args.length;
    if (argCount < 2 || argCount % 2 !== 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message:
          "IFS takes pairs of logical_test and value_if_true arguments",
        errAddress: context.dependencyNode,
      };
    }

    const evaluateArgument = (
      index: number,
      whenOmitted: FunctionEvaluationResult
    ): FunctionEvaluationResult => {
      const argument = getSuppliedArgument(node, index);
      return argument ? this.evaluateNode(argument, context) : whenOmitted;
    };
    const omittedTest: FunctionEvaluationResult = {
      type: "value",
      result: { type: "boolean", value: false },
    };
    const omittedValue: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 0 },
    };
    const noMatch: SingleEvaluationResult = {
      type: "error",
      err: FormulaError.NA,
      message: "IFS: no logical_test is TRUE",
      errAddress: context.dependencyNode,
    };

    for (let index = 0; index < argCount; index += 2) {
      const test = evaluateArgument(index, omittedTest);
      if (test.type === "awaiting-evaluation" || test.type === "error") {
        return test;
      }
      if (test.type === "value") {
        if (isTruthy(test.result)) {
          return evaluateArgument(index + 1, omittedValue);
        }
        continue;
      }

      // An array condition: the remaining pairs are needed for the shape. Their
      // cells are still read lazily, per chosen value.
      const remaining: FunctionEvaluationResult[] = [test];
      for (let rest = index + 1; rest < argCount; rest++) {
        remaining.push(
          evaluateArgument(rest, rest % 2 === 0 ? omittedTest : omittedValue)
        );
      }
      const awaiting = remaining.find(
        (result) => result.type === "awaiting-evaluation"
      );
      if (awaiting) {
        return awaiting;
      }

      return createBroadcastResult({
        operands: remaining.map((result) =>
          toBroadcastOperand(result, context.cellAddress)
        ),
        context,
        source: "IFS with an array condition",
        combine: (operandCell) => {
          for (let pair = 0; pair < remaining.length; pair += 2) {
            const cellTest = operandCell(pair);
            if (cellTest.type !== "value") {
              return cellTest;
            }
            if (isTruthy(cellTest.result)) {
              return operandCell(pair + 1);
            }
          }
          return noMatch;
        },
      });
    }

    return noMatch;
  },
};
