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
import { COUNTIF } from "./lookup/countif/countif";
import { AVERAGE } from "./math/average/average";
import { MAX } from "./math/max/max";
import { MIN } from "./math/min/min";
import { SUM } from "./math/sum/sum";
import { CEILING } from "./math/ceiling/ceiling";
import { FIND } from "./text/find/find";
import { LEFT } from "./text/left/left";
import { MID } from "./text/mid/mid";
import { LEN } from "./text/len/len";
import { RIGHT } from "./text/right/right";
import { CONCATENATE } from "./text/concatenate/concatenate";
import { IF } from "./logical/if/if";
import { IFERROR } from "./logical/iferror/iferror";

export const functions: Record<string, FunctionDefinition> = {
  AVERAGE,
  CEILING,
  CONCATENATE,
  COUNTIF,
  FIND,
  IF,
  IFERROR,
  INDEX,
  LEFT,
  LEN,
  MATCH,
  MAX,
  MID,
  MIN,
  RIGHT,
  SEQUENCE,
  SUM,
};
