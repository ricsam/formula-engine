import type { CellValue } from "../../core/types";
import type { FunctionDefinition } from "../../evaluator/evaluator";
import {
  coerceToNumber,
  coerceToString,
  isFormulaError,
  propagateError,
  validateArgCount,
} from "../index";

// Advanced mathematical functions

export const ABS: FunctionDefinition = {
  name: "ABS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ABS", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return Math.abs(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const SIGN: FunctionDefinition = {
  name: "SIGN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("SIGN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return Math.sign(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const SQRT: FunctionDefinition = {
  name: "SQRT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("SQRT", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value < 0) {
        return "#NUM!";
      }

      return Math.sqrt(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const POWER: FunctionDefinition = {
  name: "POWER",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("POWER", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const base = coerceToNumber(args[0]);
      const exponent = coerceToNumber(args[1]);

      const result = Math.pow(base, exponent);

      if (!isFinite(result)) {
        return "#NUM!";
      }

      return result;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const EXP: FunctionDefinition = {
  name: "EXP",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("EXP", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const result = Math.exp(value);

      if (!isFinite(result)) {
        return "#NUM!";
      }

      return result;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const LN: FunctionDefinition = {
  name: "LN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("LN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value <= 0) {
        return "#NUM!";
      }

      return Math.log(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const LOG: FunctionDefinition = {
  name: "LOG",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("LOG", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const base = args.length > 1 ? coerceToNumber(args[1]) : 10;

      if (value <= 0 || base <= 0 || base === 1) {
        return "#NUM!";
      }

      return Math.log(value) / Math.log(base);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const LOG10: FunctionDefinition = {
  name: "LOG10",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("LOG10", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value <= 0) {
        return "#NUM!";
      }

      return Math.log10(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const SIN: FunctionDefinition = {
  name: "SIN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("SIN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return Math.sin(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const COS: FunctionDefinition = {
  name: "COS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("COS", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return Math.cos(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const TAN: FunctionDefinition = {
  name: "TAN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("TAN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const result = Math.tan(value);

      if (!isFinite(result)) {
        return "#NUM!";
      }

      return result;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ASIN: FunctionDefinition = {
  name: "ASIN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ASIN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value < -1 || value > 1) {
        return "#NUM!";
      }

      return Math.asin(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ACOS: FunctionDefinition = {
  name: "ACOS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ACOS", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value < -1 || value > 1) {
        return "#NUM!";
      }

      return Math.acos(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ATAN: FunctionDefinition = {
  name: "ATAN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ATAN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return Math.atan(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ATAN2: FunctionDefinition = {
  name: "ATAN2",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ATAN2", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const x = coerceToNumber(args[0]);
      const y = coerceToNumber(args[1]);
      return Math.atan2(x, y);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const DEGREES: FunctionDefinition = {
  name: "DEGREES",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("DEGREES", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const radians = coerceToNumber(args[0]);
      return radians * (180 / Math.PI);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const RADIANS: FunctionDefinition = {
  name: "RADIANS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("RADIANS", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const degrees = coerceToNumber(args[0]);
      return degrees * (Math.PI / 180);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const PI: FunctionDefinition = {
  name: "PI",
  minArgs: 0,
  maxArgs: 0,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("PI", args, 0, 0);
    return Math.PI;
  },
};

export const ROUND: FunctionDefinition = {
  name: "ROUND",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ROUND", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return "#NUM!";
      }

      const factor = Math.pow(10, digits);
      return Math.round(value * factor) / factor;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ROUNDUP: FunctionDefinition = {
  name: "ROUNDUP",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ROUNDUP", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return "#NUM!";
      }

      const factor = Math.pow(10, digits);
      return value >= 0
        ? Math.ceil(value * factor) / factor
        : Math.floor(value * factor) / factor;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ROUNDDOWN: FunctionDefinition = {
  name: "ROUNDDOWN",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ROUNDDOWN", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return "#NUM!";
      }

      const factor = Math.pow(10, digits);
      return value >= 0
        ? Math.floor(value * factor) / factor
        : Math.ceil(value * factor) / factor;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const CEILING: FunctionDefinition = {
  name: "CEILING",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("CEILING", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const significance = args.length > 1 ? coerceToNumber(args[1]) : 1;

      if (significance === 0) {
        return 0;
      }

      if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) {
        return "#NUM!";
      }

      return Math.ceil(value / significance) * significance;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const FLOOR: FunctionDefinition = {
  name: "FLOOR",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FLOOR", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const significance = args.length > 1 ? coerceToNumber(args[1]) : 1;

      if (significance === 0) {
        return 0;
      }

      if ((value > 0 && significance < 0) || (value < 0 && significance > 0)) {
        return "#NUM!";
      }

      return Math.floor(value / significance) * significance;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const INT: FunctionDefinition = {
  name: "INT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("INT", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return Math.floor(value);
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const TRUNC: FunctionDefinition = {
  name: "TRUNC",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("TRUNC", args, 1, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const digits = args.length > 1 ? coerceToNumber(args[1]) : 0;

      if (!Number.isInteger(digits)) {
        return "#NUM!";
      }

      const factor = Math.pow(10, digits);
      return Math.trunc(value * factor) / factor;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const MOD: FunctionDefinition = {
  name: "MOD",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("MOD", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const dividend = coerceToNumber(args[0]);
      const divisor = coerceToNumber(args[1]);

      if (divisor === 0) {
        return "#DIV/0!";
      }

      return dividend - Math.floor(dividend / divisor) * divisor;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const EVEN: FunctionDefinition = {
  name: "EVEN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("EVEN", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value === 0) {
        return 0;
      }

      const sign = value > 0 ? 1 : -1;
      const absValue = Math.abs(value);
      const evenValue = Math.ceil(absValue / 2) * 2;

      return evenValue * sign;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const ODD: FunctionDefinition = {
  name: "ODD",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("ODD", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value === 0) {
        return 1;
      }

      const sign = value > 0 ? 1 : -1;
      const absValue = Math.abs(value);
      const oddValue = Math.ceil((absValue + 1) / 2) * 2 - 1;

      return oddValue * sign;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const FACT: FunctionDefinition = {
  name: "FACT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FACT", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);

      if (value < 0 || !Number.isInteger(value)) {
        return "#NUM!";
      }

      if (value > 170) {
        return "#NUM!"; // Factorial would overflow
      }

      let result = 1;
      for (let i = 2; i <= value; i++) {
        result *= i;
      }

      return result;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const DECIMAL: FunctionDefinition = {
  name: "DECIMAL",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("DECIMAL", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const text = coerceToString(args[0]);
      const radix = coerceToNumber(args[1]);

      if (!Number.isInteger(radix) || radix < 2 || radix > 36) {
        return "#NUM!";
      }

      const result = parseInt(text, radix);

      if (isNaN(result)) {
        return "#NUM!";
      }

      return result;
    } catch (e) {
      return "#VALUE!";
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
