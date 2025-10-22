import { type FunctionDefinition } from "src/core/types";
import { performMinimum } from "./min-utils";
import { createArgumentIterator } from "src/functions/function-utils";

/**
 * MIN function - Returns the smallest number in the arguments
 */
export const MIN: FunctionDefinition = {
  name: "MIN",
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

    // Perform minimum calculation (now with unified error propagation)
    return performMinimum(argumentIterator.values, context);
  },
};
