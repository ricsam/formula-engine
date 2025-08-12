import type { FunctionDefinition } from "src/core/types";
// import { arrayFunctions } from "./array/array-functions";
// import { infoFunctions } from "./info/info-functions";
// import { logicalComparisonFunctions } from "./logical/comparisons";
// import { logicalConditionFunctions } from "./logical/conditions";
// import { lookupFunctions } from "./lookup/lookup-functions";
// import { advancedMathFunctions } from "./math/advanced";
// import { basicMathFunctions } from "./math/basic";
// import { textFunctions } from "./text/string-functions";
import { SEQUENCE } from "./array/sequence/sequence";
import { INDEX } from "./lookup/index-lookup/index-lookup"; // Fixed import path
import { MATCH } from "./lookup/match/match";
import { SUM } from "./math/sum/sum";
import { FIND } from "./text/find/find";
import { LEFT } from "./text/left/left";
import { RIGHT } from "./text/right/right";

export const functions: Record<string, FunctionDefinition> = {
  SUM,
  LEFT,
  RIGHT,
  FIND,
  SEQUENCE,
  MATCH,
  INDEX,
};
