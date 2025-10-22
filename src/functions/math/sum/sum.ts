import { type FunctionDefinition } from "src/core/types";
import { performSummation } from "./summation-utils";
import { createArgumentIterator } from "src/functions/function-utils";

/**
 * SUM function - Adds all numbers in the arguments
 */
export const SUM: FunctionDefinition = {
  name: "SUM",
  evaluate: function (node, context) {
    // Create iterator from all arguments
    const argumentIterator = createArgumentIterator(
      this,
      node,
      context,
      "col-major"
    );

    if (argumentIterator.type !== "values") {
      return argumentIterator;
    }

    // Perform summation with error propagation (SUM propagates errors from direct arguments)
    return performSummation(argumentIterator.values);
  },
};
