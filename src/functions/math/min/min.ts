import { type FunctionDefinition } from "../../../core/types";
import { createArgumentIterator } from "../../../functions/function-utils";
import { performMinimum } from "./min-utils";

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
