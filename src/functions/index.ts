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
import { COUNT } from "./lookup/count/count";
import { COUNTIF } from "./lookup/count/countif";
import { AVERAGE } from "./math/average/average";
import { AVERAGEIF } from "./math/average/averageif";
import { AVERAGEIFS } from "./math/average/averageifs";
import { MAX } from "./math/max/max";
import { MAXIFS } from "./math/max/maxifs";
import { MIN } from "./math/min/min";
import { MINIFS } from "./math/min/minifs";
import { SUM } from "./math/sum/sum";
import { SUMIF } from "./math/sum/sumif";
import { SUMIFS } from "./math/sum/sumifs";
import { COUNTIFS } from "./lookup/count/countifs";
import { MAXIF } from "./math/max/maxif";
import { MINIF } from "./math/min/minif";
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
import { XLOOKUP } from "./lookup/xlookup/xlookup";
import { OR } from "./logical/or/or";
import { TEXTJOIN } from "./text/textjoin/textjoin";
import { ROW } from "./information/row/row";
import { COLUMN } from "./information/column/column";
import { CELL } from "./information/cell/cell";
import { ADDRESS } from "./reference/address/address";
import { INDIRECT } from "./reference/indirect/indirect";
import { OFFSET } from "./reference/offset/offset";

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
    ADDRESS,
    AND,
    AVERAGE,
    AVERAGEIF,
    AVERAGEIFS,
    CEILING,
    CELL,
    COLUMN,
    CONCATENATE,
    COUNT,
    COUNTIF,
    COUNTIFS,
    EXACT,
    FIND,
    IF,
    IFERROR,
    INDEX,
    INDIRECT,
    LEFT,
    LEN,
    MATCH,
    MAX,
    MAXIF,
    MAXIFS,
    MID,
    MIN,
    MINIF,
    MINIFS,
    OFFSET,
    OR,
    RIGHT,
    ROW,
    SEQUENCE,
    SUM,
    SUMIF,
    SUMIFS,
    TEXTJOIN,
    XLOOKUP,
  }
);
