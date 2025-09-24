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
import { COUNTIF } from "./lookup/count/countif";
import { AVERAGE } from "./math/average/average";
import { AVERAGEIF } from "./math/average/averageif";
import { AVERAGEIFS } from "./math/average/averageifs";
import { MAX } from "./math/max/max";
import { MAXIFS } from "./math/max/maxifs";
import { MIN } from "./math/min/min";
import { MINIFS } from "./math/min/minifs";
import { SUM } from "./math/sum/sum";
import { CEILING } from "./math/ceiling/ceiling";
import { EXACT } from "./text/exact/exact";
import { FIND } from "./text/find/find";
import { LEFT } from "./text/left/left";
import { MID } from "./text/mid/mid";
import { LEN } from "./text/len/len";
import { RIGHT } from "./text/right/right";
import { CONCATENATE } from "./text/concatenate/concatenate";
import { AND } from "./logical/and/and";
import { IF } from "./logical/if/if";
import { IFERROR } from "./logical/iferror/iferror";

const buildFunctionIndex = (functions: Record<string, FunctionDefinition>) => {
  return Object.fromEntries(
    Object.entries(functions).flatMap(([name, func]) => {
      const base: [string, FunctionDefinition][] = [[name, func]];
      if (func.aliases) {
        func.aliases.forEach((alias) => {
          base.push([alias, func]);
        });
      }
      return base;
    })
  );
};

export const functions: Record<string, FunctionDefinition> = buildFunctionIndex(
  {
    AND,
    AVERAGE,
    AVERAGEIF,
    AVERAGEIFS,
    CEILING,
    CONCATENATE,
    COUNTIF,
    EXACT,
    FIND,
    IF,
    IFERROR,
    INDEX,
    LEFT,
    LEN,
    MATCH,
    MAX,
    MAXIFS,
    MID,
    MIN,
    MINIFS,
    RIGHT,
    SEQUENCE,
    SUM,
  }
);
