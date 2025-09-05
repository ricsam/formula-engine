import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type CellInfinity,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
  type ErrorEvaluationResult,
} from "src/core/types";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * CEILING function - Rounds a number up to the nearest multiple of significance
 * 
 * Usage: CEILING(number, significance)
 * 
 * number: The number to round up (required)
 * significance: The multiple to round up to (required)
 * 
 * Examples:
 *   CEILING(4.3) returns 5 (rounds up to nearest integer)
 *   CEILING(4.3, 0.5) returns 4.5 (rounds up to nearest 0.5)
 *   CEILING(-2.1, -1) returns -3 (rounds away from zero)
 * 
 * Note:
 * - Both arguments are required and must be numbers (no type coercion)
 * - Positive number with negative significance returns #NUM! error
 * - Supports spilled values (dynamic arrays)
 * - Handles infinity values appropriately
 */

/**
 * Helper function to perform CEILING operation
 */
function ceilingOperation(
  numberValue: CellNumber | CellInfinity,
  significanceValue: CellNumber | CellInfinity
): CellNumber | CellInfinity | ErrorEvaluationResult {
  // Handle infinity cases
  if (numberValue.type === "infinity") {
    return numberValue; // Infinity remains infinity
  }
  
  if (significanceValue.type === "infinity") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Cannot use infinity as significance",
    };
  }
  
  // Both are numbers
  const num = numberValue.value;
  const sig = significanceValue.value;
  
  // Significance cannot be zero
  if (sig === 0) {
    return {
      type: "error",
      err: FormulaError.DIV0,
      message: "Significance cannot be zero",
    };
  }
  
  // Excel CEILING behavior based on testing:
  // - Always rounds away from zero
  // - Positive number with negative significance gives #NUM! error
  
  if (num > 0 && sig < 0) {
    return {
      type: "error",
      err: FormulaError.NUM,
      message: "Positive number cannot be used with negative significance",
    };
  }
  
  if (num < 0 && sig > 0) {
    return {
      type: "error",
      err: FormulaError.NUM,
      message: "Negative number cannot be used with positive significance",
    };
  }
  
  // Calculate result - always rounds away from zero
  let result: number;
  
  if (sig > 0) {
    // Positive significance with positive/zero number
    result = Math.ceil(num / sig) * sig;
  } else {
    // Negative significance with negative/zero number
    // Round away from zero: -2.1/-1 = 2.1, ceil(2.1) = 3, 3*-1 = -3
    result = Math.ceil(num / sig) * sig;
  }
  
  return { type: "number", value: result };
}

/**
 * Helper for creating spilled-values result for CEILING function
 */
function createCeilingSpilledResult(
  this: FormulaEvaluator,
  {
    numberResult,
    significanceResult,
    context,
  }: {
    numberResult: FunctionEvaluationResult;
    significanceResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): SpilledValuesEvaluationResult | ErrorEvaluationResult {
  const hasSpilledNumber = numberResult.type === "spilled-values";
  const hasSpilledSignificance = significanceResult.type === "spilled-values";

  if (!hasSpilledNumber && !hasSpilledSignificance) {
    throw new Error("createCeilingSpilledResult called without spilled values");
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress): SpreadsheetRange => {
      // Calculate spill area (union of spilled ranges)
      if (hasSpilledNumber && numberResult.type === "spilled-values") {
        const spillArea = numberResult.spillArea(origin);
        if (hasSpilledSignificance && significanceResult.type === "spilled-values") {
          return this.unionRanges(
            this.projectRange(spillArea, origin),
            this.projectRange(significanceResult.spillArea(origin), origin)
          );
        }
        return spillArea;
      } else if (hasSpilledSignificance && significanceResult.type === "spilled-values") {
        return significanceResult.spillArea(origin);
      } else {
        throw new Error("No spilled values found");
      }
    },
    source: "CEILING with spilled values",
    evaluate: (spilledCell: any, evalContext: any): SingleEvaluationResult => {
      // Evaluate arguments at this spilled position
      const spillNumberResult = hasSpilledNumber
        ? numberResult.evaluate(spilledCell, evalContext)
        : numberResult;
      const spillSigResult = hasSpilledSignificance
        ? significanceResult.evaluate(spilledCell, evalContext)
        : significanceResult;

      if (spillNumberResult === undefined || spillSigResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled results have not been evaluated",
        };
      }
      
      if (spillNumberResult.type === "error") {
        return spillNumberResult;
      }
      if (spillSigResult.type === "error") {
        return spillSigResult;
      }

      if (spillNumberResult.type !== "value" || spillSigResult.type !== "value") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid spilled result for CEILING",
        };
      }

      // Type checking
      if (spillNumberResult.result.type !== "number" && spillNumberResult.result.type !== "infinity") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Number argument must be a number",
        };
      }
      
      if (spillSigResult.result.type !== "number" && spillSigResult.result.type !== "infinity") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Significance argument must be a number",
        };
      }

      const result = ceilingOperation(spillNumberResult.result, spillSigResult.result);
      if (result.type === "error") {
        return result;
      }
      return { type: "value", result };
    },
    evaluateAllCells: (intersectingRange: any) => {
      throw new Error("WIP: evaluateAllCells for CEILING is not implemented");
    },
  };
}

/**
 * CEILING function implementation
 */
export const CEILING: FunctionDefinition = {
  name: "CEILING",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "CEILING function takes exactly 2 arguments",
      };
    }

    // Evaluate number argument
    const numberResult = this.evaluateNode(node.args[0]!, context);
    if (numberResult.type === "error") {
      return numberResult;
    }

    // Evaluate significance argument (required)
    const significanceResult = this.evaluateNode(node.args[1]!, context);
    if (significanceResult.type === "error") {
      return significanceResult;
    }

    // Handle spilled values
    const hasSpilledValues = 
      numberResult.type === "spilled-values" ||
      significanceResult.type === "spilled-values";

    if (hasSpilledValues) {
      return createCeilingSpilledResult.call(this, {
        numberResult,
        significanceResult,
        context,
      });
    }

    // Both arguments are single values
    if (numberResult.type !== "value" || significanceResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid argument type",
      };
    }

    // Type checking - only numbers and infinity allowed
    if (numberResult.result.type !== "number" && numberResult.result.type !== "infinity") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Number argument must be a number",
      };
    }
    
    if (significanceResult.result.type !== "number" && significanceResult.result.type !== "infinity") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Significance argument must be a number",
      };
    }

    // Perform CEILING operation
    const result = ceilingOperation(numberResult.result, significanceResult.result);
    if (result.type === "error") {
      return result;
    }
    return { type: "value", result };
  },
};
