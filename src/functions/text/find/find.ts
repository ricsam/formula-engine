import {
  FormulaError,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";
import { convertToString, extractNumericValue } from "../text-helpers";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

// Helper function for FIND operation - returns the result or null if error
function findOperation(
  findTextResult: FunctionEvaluationResult,
  withinTextResult: FunctionEvaluationResult,
  startNumResult: FunctionEvaluationResult
): CellNumber | null {
  try {
    const findStr = convertToString(findTextResult);
    const withinStr = convertToString(withinTextResult);
    const startNum = extractNumericValue(startNumResult);

    // Validate startNum
    if (startNum < 1) {
      return null;
    }

    const start = Math.floor(startNum) - 1; // Convert to 0-based index

    if (start >= withinStr.length) {
      return null;
    }

    const index = withinStr.indexOf(findStr, start);
    if (index === -1) {
      return null;
    }

    return { type: "number", value: index + 1 }; // Convert back to 1-based index
  } catch (error) {
    return null;
  }
}

/**
 * Helper for creating spilled-values result for FIND function
 */
function createFindSpilledResult(
  this: FormulaEvaluator,
  {
    findTextResult,
    withinTextResult,
    startNumResult,
    context,
  }: {
    findTextResult: FunctionEvaluationResult;
    withinTextResult: FunctionEvaluationResult;
    startNumResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): FunctionEvaluationResult {
  // If both findText and withinText are spilled-values, we need to zip them together
  if (
    findTextResult.type === "spilled-values" &&
    withinTextResult.type === "spilled-values"
  ) {
    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => {
        // Create unified spill area (union of all spilled ranges)
        let spillArea = this.unionRanges(
          this.projectRange(findTextResult.spillArea(origin), origin),
          this.projectRange(withinTextResult.spillArea(origin), origin)
        );

        // Also include startNum spill area if it's spilled
        if (startNumResult.type === "spilled-values") {
          spillArea = this.unionRanges(
            spillArea,
            this.projectRange(startNumResult.spillArea(origin), origin)
          );
        }
        return spillArea;
      },
      source: "FIND with zipped spilled findText and withinText values",
      evaluate: (spillOffset, evalContext: EvaluationContext) => {
        // Evaluate both spilled arrays at this position
        const spillFindResult = findTextResult.evaluate(
          spillOffset,
          evalContext
        );
        const spillWithinResult = withinTextResult.evaluate(
          spillOffset,
          evalContext
        );

        if (!spillFindResult || spillFindResult.type === "error") {
          return spillFindResult;
        }
        if (!spillWithinResult || spillWithinResult.type === "error") {
          return spillWithinResult;
        }

        // Handle startNum - evaluate if spilled, otherwise use as-is
        let startNumArg = startNumResult;
        if (startNumResult.type === "spilled-values") {
          const spillStartNumResult = startNumResult.evaluate(
            spillOffset,
            evalContext
          );
          if (!spillStartNumResult || spillStartNumResult.type === "error") {
            return spillStartNumResult;
          }
          startNumArg = spillStartNumResult;
        }

        const result = findOperation(
          spillFindResult,
          spillWithinResult,
          startNumArg
        );
        if (result === null) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Text not found #2",
          };
        }
        return {
          type: "value",
          result: result,
        };
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for FIND is not implemented");
      },
    };
  }
  // If only findText is spilled-values
  else if (
    findTextResult.type === "spilled-values" &&
    withinTextResult.type !== "spilled-values" &&
    startNumResult.type !== "spilled-values"
  ) {
    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => {
        return findTextResult.spillArea(origin);
      },
      source: "FIND with spilled findText values",
      evaluate: (spillOffset, evalContext: EvaluationContext) => {
        const spillResult = findTextResult.evaluate(spillOffset, evalContext);
        if (!spillResult || spillResult.type === "error") {
          return spillResult;
        }
        const result = findOperation(
          spillResult,
          withinTextResult,
          startNumResult
        );
        if (result === null) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Text not found #4",
          };
        }
        return {
          type: "value",
          result: result,
        };
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for FIND is not implemented");
      },
    };
  }
  // If only withinText is spilled-values
  else if (
    withinTextResult.type === "spilled-values" &&
    findTextResult.type !== "spilled-values" &&
    startNumResult.type !== "spilled-values"
  ) {
    return {
      type: "spilled-values",
      spillArea: (origin) => withinTextResult.spillArea(origin),
      source: "FIND with spilled withinText values",
      evaluate: (spillOffset, evalContext: EvaluationContext) => {
        const spillResult = withinTextResult.evaluate(spillOffset, evalContext);
        if (!spillResult || spillResult.type === "error") {
          return spillResult;
        }
        const result = findOperation(
          findTextResult,
          spillResult,
          startNumResult
        );
        if (result === null) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Text not found #6",
          };
        }
        return {
          type: "value",
          result: result,
        };
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for FIND is not implemented");
      },
    };
  }
  // If findText and startNum are spilled (but not withinText)
  else if (
    findTextResult.type === "spilled-values" &&
    startNumResult.type === "spilled-values" &&
    withinTextResult.type !== "spilled-values"
  ) {
    if (withinTextResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid withinText argument",
      };
    }

    // Strict type checking
    if (withinTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "WithinText argument must be a string",
      };
    }

    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => {
        // Create unified spill area (union of both ranges)
        const spillArea = this.unionRanges(
          this.projectRange(findTextResult.spillArea(origin), origin),
          this.projectRange(startNumResult.spillArea(origin), origin)
        );
        return spillArea;
      },
      source: "FIND with spilled findText and startNum values",
      evaluate: (spillOffset, evalContext: EvaluationContext) => {
        // Evaluate both spilled arrays at this position
        const spillFindResult = findTextResult.evaluate(
          spillOffset,
          evalContext
        );
        const spillStartNumResult = startNumResult.evaluate(
          spillOffset,
          evalContext
        );

        if (!spillFindResult || spillFindResult.type === "error") {
          return spillFindResult;
        }
        if (!spillStartNumResult || spillStartNumResult.type === "error") {
          return spillStartNumResult;
        }

        const result = findOperation(
          spillFindResult,
          withinTextResult,
          spillStartNumResult
        );
        if (result === null) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Text not found #8",
          };
        }
        return {
          type: "value",
          result: result,
        };
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for FIND is not implemented");
      },
    };
  }
  // If withinText and startNum are spilled (but not findText)
  else if (
    withinTextResult.type === "spilled-values" &&
    startNumResult.type === "spilled-values" &&
    findTextResult.type !== "spilled-values"
  ) {
    if (findTextResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid findText argument",
      };
    }

    // Strict type checking
    if (findTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "FindText argument must be a string",
      };
    }


    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => {
        // Create unified spill area (union of both ranges)
        const spillArea = this.unionRanges(
          this.projectRange(withinTextResult.spillArea(origin), origin),
          this.projectRange(startNumResult.spillArea(origin), origin)
        );
        return spillArea;
      },
      source: "FIND with spilled withinText and startNum values",
      evaluate: (spillOffset, evalContext: EvaluationContext) => {
        // Evaluate both spilled arrays at this position
        const spillWithinResult = withinTextResult.evaluate(
          spillOffset,
          evalContext
        );
        const spillStartNumResult = startNumResult.evaluate(
          spillOffset,
          evalContext
        );

        if (!spillWithinResult || spillWithinResult.type === "error") {
          return spillWithinResult;
        }
        if (!spillStartNumResult || spillStartNumResult.type === "error") {
          return spillStartNumResult;
        }

        const result = findOperation(
          findTextResult,
          spillWithinResult,
          spillStartNumResult
        );
        if (result === null) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Text not found #10",
          };
        }
        return {
          type: "value",
          result: result,
        };
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for FIND is not implemented");
      },
    };
  }
  // If only startNum is spilled-values
  else if (
    startNumResult.type === "spilled-values" &&
    findTextResult.type !== "spilled-values" &&
    withinTextResult.type !== "spilled-values"
  ) {
    if (findTextResult.type !== "value" || withinTextResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid findText or withinText argument",
      };
    }

    // Strict type checking
    if (findTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "FindText argument must be a string",
      };
    }

    if (withinTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "WithinText argument must be a string",
      };
    }

    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => startNumResult.spillArea(origin),
      source: "FIND with spilled startNum values",
      evaluate: (spillOffset, evalContext: EvaluationContext) => {
        const spillResult = startNumResult.evaluate(spillOffset, evalContext);
        if (!spillResult || spillResult.type === "error") {
          return spillResult;
        }

        // Create a proper startNum argument for findOperation
        if (spillResult.type !== "value") {
          return spillResult;
        }
        const startNumArg: ValueEvaluationResult = {
          type: "value",
          result: spillResult.result,
        };

        const result = findOperation(
          findTextResult,
          withinTextResult,
          startNumArg
        );
        if (result === null) {
          return {
            type: "error",
            err: FormulaError.VALUE,
            message: "Text not found #12",
          };
        }
        return {
          type: "value",
          result: result,
        };
      },
      evaluateAllCells: (intersectingRange) => {
        throw new Error("WIP: evaluateAllCells for FIND is not implemented");
      },
    };
  }
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Invalid arguments for FIND",
  };
}

/**
 * FIND function - Finds one text string within another (case-sensitive)
 *
 * Usage: FIND(find_text, within_text, [start_num])
 *
 * find_text: The text you want to find.
 * within_text: The text containing the text you want to find.
 * start_num: [Optional] The character at which to start the search (default: 1).
 *
 * Example: FIND("World", "Hello World", 1) returns 7
 *
 * Note:
 * - The search is case-sensitive
 * - Returns the position of the first character of find_text within within_text
 * - Returns #VALUE! error if text is not found
 * - Supports dynamic arrays (spilled values) for find_text and within_text
 * - Strict type checking: text arguments must be strings, start_num must be number
 */
export const FIND: FunctionDefinition = {
  name: "FIND",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "FIND function takes 2 or 3 arguments",
      };
    }

    // Evaluate findText argument
    const findTextResult = this.evaluateNode(node.args[0]!, context);
    if (findTextResult.type === "error") {
      return findTextResult;
    }

    // Evaluate withinText argument
    const withinTextResult = this.evaluateNode(node.args[1]!, context);
    if (withinTextResult.type === "error") {
      return withinTextResult;
    }

    // Evaluate startNum argument (optional, defaults to 1)
    let startNumResult: FunctionEvaluationResult;
    if (node.args.length > 2) {
      startNumResult = this.evaluateNode(node.args[2]!, context);
    } else {
      startNumResult = {
        type: "value",
        result: { type: "number", value: 1 },
      };
    }

    // Handle spilled-values inputs - return spilled-values for spilling
    if (
      findTextResult.type === "spilled-values" ||
      withinTextResult.type === "spilled-values" ||
      startNumResult.type === "spilled-values"
    ) {
      return createFindSpilledResult.call(this, {
        findTextResult,
        withinTextResult,
        startNumResult,
        context,
      });
    }

    // Both findText and withinText are single values
    if (findTextResult.type !== "value" || withinTextResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid findText or withinText argument",
      };
    }

    // Strict type checking - no coercion
    if (findTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "FindText argument must be a string",
      };
    }

    if (withinTextResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "WithinText argument must be a string",
      };
    }

    const result = findOperation(
      findTextResult,
      withinTextResult,
      startNumResult
    );

    if (result === null) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Text not found #13",
      };
    }

    return {
      type: "value",
      result: result,
    };
  },
};
