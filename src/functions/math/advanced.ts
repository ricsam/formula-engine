import { FormulaError, type CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import {
  coerceToNumber,
  coerceToString,
  isFormulaError,
  propagateError,
  validateArgCount,
} from "../utils";

// Advanced mathematical functions

const ABS: FunctionDefinition = {
  name: "ABS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ABS", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: Math.abs(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const SIGN: FunctionDefinition = {
  name: "SIGN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("SIGN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: Math.sign(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const SQRT: FunctionDefinition = {
  name: "SQRT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("SQRT", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value < 0) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: Math.sqrt(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const POWER: FunctionDefinition = {
  name: "POWER",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("POWER", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const base = coerceToNumber(args[0]);
      const exponent = coerceToNumber(args[1]);

      const result = Math.pow(base, exponent);

      if (!isFinite(result)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const EXP: FunctionDefinition = {
  name: "EXP",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("EXP", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const result = Math.exp(value);

      if (!isFinite(result)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const LN: FunctionDefinition = {
  name: "LN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("LN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value <= 0) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: Math.log(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const LOG: FunctionDefinition = {
  name: "LOG",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("LOG", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const base = args.length > 1 ? coerceToNumber(args[1]) : 10;

      if (value <= 0 || base <= 0 || base === 1) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: Math.log(value) / Math.log(base) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const LOG10: FunctionDefinition = {
  name: "LOG10",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("LOG10", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value <= 0) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: Math.log10(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const SIN: FunctionDefinition = {
  name: "SIN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("SIN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: Math.sin(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const COS: FunctionDefinition = {
  name: "COS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("COS", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: Math.cos(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const TAN: FunctionDefinition = {
  name: "TAN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("TAN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const result = Math.tan(value);

      if (!isFinite(result)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ASIN: FunctionDefinition = {
  name: "ASIN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ASIN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value < -1 || value > 1) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: Math.asin(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ACOS: FunctionDefinition = {
  name: "ACOS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ACOS", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value < -1 || value > 1) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: Math.acos(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ATAN: FunctionDefinition = {
  name: "ATAN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ATAN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: Math.atan(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ATAN2: FunctionDefinition = {
  name: "ATAN2",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ATAN2", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const x = coerceToNumber(args[0]);
      const y = coerceToNumber(args[1]);
      return { type: "value", value: Math.atan2(x, y) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const DEGREES: FunctionDefinition = {
  name: "DEGREES",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("DEGREES", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const radians = coerceToNumber(args[0]);
      return { type: "value", value: radians * (180 / Math.PI) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const RADIANS: FunctionDefinition = {
  name: "RADIANS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("RADIANS", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const degrees = coerceToNumber(args[0]);
      return { type: "value", value: degrees * (Math.PI / 180) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const PI: FunctionDefinition = {
  name: "PI",
  minArgs: 0,
  maxArgs: 0,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("PI", args, 0, 0);
    return { type: "value", value: Math.PI };
  },
};

const ROUND: FunctionDefinition = {
  name: "ROUND",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ROUND", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return { type: "value", value: FormulaError.NUM };
      }

      const factor = Math.pow(10, digits);
      return { type: "value", value: Math.round(value * factor) / factor };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ROUNDUP: FunctionDefinition = {
  name: "ROUNDUP",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ROUNDUP", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return { type: "value", value: FormulaError.NUM };
      }

      const factor = Math.pow(10, digits);
      return {
        type: "value",
        value:
          value >= 0
            ? Math.ceil(value * factor) / factor
            : Math.floor(value * factor) / factor,
      };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ROUNDDOWN: FunctionDefinition = {
  name: "ROUNDDOWN",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ROUNDDOWN", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return { type: "value", value: FormulaError.NUM };
      }

      const factor = Math.pow(10, digits);
      return {
        type: "value",
        value:
          value >= 0
            ? Math.floor(value * factor) / factor
            : Math.ceil(value * factor) / factor,
      };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const CEILING: FunctionDefinition = {
  name: "CEILING",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("CEILING", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const significance = args.length > 1 ? coerceToNumber(args[1]) : 1;

      if (significance === 0) {
        return { type: "value", value: 0 };
      }

      if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return {
        type: "value",
        value: Math.ceil(value / significance) * significance,
      };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const FLOOR: FunctionDefinition = {
  name: "FLOOR",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FLOOR", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const significance = args.length > 1 ? coerceToNumber(args[1]) : 1;

      if (significance === 0) {
        return { type: "value", value: 0 };
      }

      if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return {
        type: "value",
        value: Math.floor(value / significance) * significance,
      };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const INT: FunctionDefinition = {
  name: "INT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("INT", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: Math.floor(value) };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const TRUNC: FunctionDefinition = {
  name: "TRUNC",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("TRUNC", args, 1, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return { type: "value", value: FormulaError.NUM };
      }

      const factor = Math.pow(10, digits);
      return { type: "value", value: Math.trunc(value * factor) / factor };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const MOD: FunctionDefinition = {
  name: "MOD",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("MOD", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const dividend = coerceToNumber(args[0]);
      const divisor = coerceToNumber(args[1]);

      if (divisor === 0) {
        return { type: "value", value: FormulaError.DIV0 };
      }

      return {
        type: "value",
        value: dividend - Math.floor(dividend / divisor) * divisor,
      };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const EVEN: FunctionDefinition = {
  name: "EVEN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("EVEN", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value === 0) {
        return { type: "value", value: 0 };
      }

      const sign = value > 0 ? 1 : -1;
      const absValue = Math.abs(value);
      const evenValue = Math.ceil(absValue / 2) * 2;

      return { type: "value", value: evenValue * sign };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const ODD: FunctionDefinition = {
  name: "ODD",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("ODD", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value === 0) {
        return { type: "value", value: 1 };
      }

      const sign = value > 0 ? 1 : -1;
      const absValue = Math.abs(value);
      const oddValue = Math.ceil((absValue + 1) / 2) * 2 - 1;

      return { type: "value", value: oddValue * sign };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const FACT: FunctionDefinition = {
  name: "FACT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FACT", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);

      if (value < 0 || !Number.isInteger(value)) {
        return { type: "value", value: FormulaError.NUM };
      }

      if (value > 170) {
        return { type: "value", value: FormulaError.NUM }; // Factorial would overflow
      }

      let result = 1;
      for (let i = 2; i <= value; i++) {
        result *= i;
      }

      return { type: "value", value: result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const DECIMAL: FunctionDefinition = {
  name: "DECIMAL",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("DECIMAL", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const text = coerceToString(args[0]);
      const radix = coerceToNumber(args[1]);

      if (!Number.isInteger(radix) || radix < 2 || radix > 36) {
        return { type: "value", value: FormulaError.NUM };
      }

      const result = parseInt(text, radix);

      if (isNaN(result)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

// Export all advanced math functions
export const advancedMathFunctions: FunctionDefinition[] = [
  ABS,
  SIGN,
  SQRT,
  POWER,
  EXP,
  LN,
  LOG,
  LOG10,
  SIN,
  COS,
  TAN,
  ASIN,
  ACOS,
  ATAN,
  ATAN2,
  DEGREES,
  RADIANS,
  PI,
  ROUND,
  ROUNDUP,
  ROUNDDOWN,
  CEILING,
  FLOOR,
  INT,
  TRUNC,
  MOD,
  EVEN,
  ODD,
  FACT,
  DECIMAL,
];
