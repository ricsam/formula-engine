import { FormulaError, type CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import { isFormulaError, propagateError } from "../utils";

/**
 * CONCATENATE function - Concatenates multiple text values
 */
const CONCATENATE: FunctionDefinition = {
  name: "CONCATENATE",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    let result = "";
    for (const arg of args) {
      if (arg === undefined || arg === null) {
        // Skip undefined/null values
        continue;
      }
      result += String(arg);
    }

    return { type: "value", value: result };
  },
};

/**
 * LEN function - Returns the length of a text string
 */
const LEN: FunctionDefinition = {
  name: "LEN",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const value = args[0];
    if (value === undefined || value === null) {
      return { type: "value", value: 0 };
    }

    return { type: "value", value: String(value).length };
  },
};

/**
 * UPPER function - Converts text to uppercase
 */
const UPPER: FunctionDefinition = {
  name: "UPPER",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const value = args[0];
    if (value === undefined || value === null) {
      return { type: "value", value: "" };
    }

    return { type: "value", value: String(value).toUpperCase() };
  },
};

/**
 * LOWER function - Converts text to lowercase
 */
const LOWER: FunctionDefinition = {
  name: "LOWER",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const value = args[0];
    if (value === undefined || value === null) {
      return { type: "value", value: "" };
    }

    return { type: "value", value: String(value).toLowerCase() };
  },
};

/**
 * TRIM function - Removes extra spaces from text
 */
const TRIM: FunctionDefinition = {
  name: "TRIM",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const value = args[0];
    if (value === undefined || value === null) {
      return { type: "value", value: "" };
    }

    return { type: "value", value: String(value).trim().replace(/\s+/g, " ") };
  },
};

/**
 * LEFT function - Returns the leftmost characters from a text string
 */
const LEFT: FunctionDefinition = {
  name: "LEFT",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const text = args[0];
    const numChars = args.length > 1 ? args[1] : 1;

    if (text === undefined || text === null) {
      return { type: "value", value: "" };
    }

    if (typeof numChars !== "number" || numChars < 0) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const textStr = String(text);
    return { type: "value", value: textStr.substring(0, Math.floor(numChars)) };
  },
};

/**
 * RIGHT function - Returns the rightmost characters from a text string
 */
const RIGHT: FunctionDefinition = {
  name: "RIGHT",
  minArgs: 1,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const text = args[0];
    const numChars = args.length > 1 ? args[1] : 1;

    if (text === undefined || text === null) {
      return { type: "value", value: "" };
    }

    if (typeof numChars !== "number" || numChars < 0) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const textStr = String(text);
    const start = Math.max(0, textStr.length - Math.floor(numChars));
    return { type: "value", value: textStr.substring(start) };
  },
};

/**
 * MID function - Returns characters from the middle of a text string
 */
const MID: FunctionDefinition = {
  name: "MID",
  minArgs: 3,
  maxArgs: 3,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const text = args[0];
    const startNum = args[1];
    const numChars = args[2];

    if (text === undefined || text === null) {
      return { type: "value", value: "" };
    }

    if (
      typeof startNum !== "number" ||
      typeof numChars !== "number" ||
      startNum < 1 ||
      numChars < 0
    ) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const textStr = String(text);
    const start = Math.floor(startNum) - 1; // Convert to 0-based index
    const length = Math.floor(numChars);

    if (start >= textStr.length) {
      return { type: "value", value: "" };
    }

    return { type: "value", value: textStr.substring(start, start + length) };
  },
};

/**
 * FIND function - Finds one text string within another (case-sensitive)
 */
const FIND: FunctionDefinition = {
  name: "FIND",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const findText = args[0];
    const withinText = args[1];
    const startNum = args.length > 2 ? args[2] : 1;

    if (
      findText === undefined ||
      findText === null ||
      withinText === undefined ||
      withinText === null
    ) {
      return { type: "value", value: FormulaError.VALUE };
    }

    if (typeof startNum !== "number" || startNum < 1) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const findStr = String(findText);
    const withinStr = String(withinText);
    const start = Math.floor(startNum) - 1; // Convert to 0-based index

    if (start >= withinStr.length) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const index = withinStr.indexOf(findStr, start);
    return {
      type: "value",
      value: index === -1 ? FormulaError.VALUE : index + 1,
    }; // Convert back to 1-based index
  },
};

/**
 * SEARCH function - Finds one text string within another (case-insensitive, supports wildcards)
 */
const SEARCH: FunctionDefinition = {
  name: "SEARCH",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const findText = args[0];
    const withinText = args[1];
    const startNum = args.length > 2 ? args[2] : 1;

    if (
      findText === undefined ||
      findText === null ||
      withinText === undefined ||
      withinText === null
    ) {
      return { type: "value", value: FormulaError.VALUE };
    }

    if (typeof startNum !== "number" || startNum < 1) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const findStr = String(findText).toLowerCase();
    const withinStr = String(withinText).toLowerCase();
    const start = Math.floor(startNum) - 1; // Convert to 0-based index

    if (start >= withinStr.length) {
      return { type: "value", value: FormulaError.VALUE };
    }

    // Handle wildcards
    let pattern = findStr;
    if (findStr.includes("*") || findStr.includes("?")) {
      pattern = findStr
        .replace(/[.*+?^${}()|[\]\\]/g, "\\$&") // Escape regex special chars
        .replace(/\\\*/g, ".*") // Replace \* with .*
        .replace(/\\\?/g, "."); // Replace \? with .

      const regex = new RegExp(pattern);
      const match = withinStr.substring(start).match(regex);
      if (match) {
        return { type: "value", value: start + match.index! + 1 }; // Convert back to 1-based index
      } else {
        return { type: "value", value: FormulaError.VALUE };
      }
    } else {
      const index = withinStr.indexOf(findStr, start);
      return {
        type: "value",
        value: index === -1 ? FormulaError.VALUE : index + 1,
      }; // Convert back to 1-based index
    }
  },
};

/**
 * SUBSTITUTE function - Substitutes new text for old text in a text string
 */
const SUBSTITUTE: FunctionDefinition = {
  name: "SUBSTITUTE",
  minArgs: 3,
  maxArgs: 4,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const text = args[0];
    const oldText = args[1];
    const newText = args[2];
    const instanceNum = args.length > 3 ? args[3] : undefined;

    if (
      text === undefined ||
      text === null ||
      oldText === undefined ||
      oldText === null ||
      newText === undefined ||
      newText === null
    ) {
      return { type: "value", value: String(text || "") };
    }

    if (
      instanceNum !== undefined &&
      (typeof instanceNum !== "number" || instanceNum < 1)
    ) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const textStr = String(text);
    const oldStr = String(oldText);
    const newStr = String(newText);

    if (oldStr === "") {
      return { type: "value", value: textStr }; // Can't substitute empty string
    }

    if (instanceNum === undefined) {
      // Replace all instances
      return { type: "value", value: textStr.replaceAll(oldStr, newStr) };
    } else {
      // Replace specific instance
      const instance = Math.floor(instanceNum);
      let count = 0;
      let result = textStr;
      let index = 0;

      while ((index = result.indexOf(oldStr, index)) !== -1) {
        count++;
        if (count === instance) {
          result =
            result.substring(0, index) +
            newStr +
            result.substring(index + oldStr.length);
          break;
        }
        index += oldStr.length;
      }

      return { type: "value", value: result };
    }
  },
};

/**
 * REPLACE function - Replaces part of a text string with different text
 */
const REPLACE: FunctionDefinition = {
  name: "REPLACE",
  minArgs: 4,
  maxArgs: 4,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const oldText = args[0];
    const startNum = args[1];
    const numChars = args[2];
    const newText = args[3];

    if (
      oldText === undefined ||
      oldText === null ||
      newText === undefined ||
      newText === null
    ) {
      return { type: "value", value: FormulaError.VALUE };
    }

    if (
      typeof startNum !== "number" ||
      typeof numChars !== "number" ||
      startNum < 1 ||
      numChars < 0
    ) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const oldStr = String(oldText);
    const newStr = String(newText);
    const start = Math.floor(startNum) - 1; // Convert to 0-based index
    const length = Math.floor(numChars);

    if (start >= oldStr.length) {
      return { type: "value", value: oldStr + newStr };
    }

    return {
      type: "value",
      value:
        oldStr.substring(0, start) + newStr + oldStr.substring(start + length),
    };
  },
};

/**
 * FE.CONCAT function - Binary concatenation (simpler than CONCATENATE)
 */
const FE_CONCAT: FunctionDefinition = {
  name: "FE.CONCAT",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const text1 = args[0];
    const text2 = args[1];

    // Convert to strings, treating null/undefined as empty string
    const str1 = text1 === undefined || text1 === null ? "" : String(text1);
    const str2 = text2 === undefined || text2 === null ? "" : String(text2);

    return { type: "value", value: str1 + str2 };
  },
};

/**
 * EXACT function - Exact string comparison (case-sensitive)
 */
const EXACT: FunctionDefinition = {
  name: "EXACT",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const text1 = args[0];
    const text2 = args[1];

    // Convert to strings, treating null/undefined as empty string
    const str1 = text1 === undefined || text1 === null ? "" : String(text1);
    const str2 = text2 === undefined || text2 === null ? "" : String(text2);

    return { type: "value", value: str1 === str2 };
  },
};

/**
 * TEXT function - Format number as text (basic implementation)
 */
const TEXT: FunctionDefinition = {
  name: "TEXT",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const value = args[0];
    const format = args[1];

    if (format === undefined || format === null) {
      return { type: "value", value: FormulaError.VALUE };
    }

    const formatStr = String(format);

    // Handle number formatting
    if (typeof value === "number") {
      // Basic format handling - more complex formatting could be added later
      if (formatStr === "0") {
        return { type: "value", value: Math.round(value).toString() };
      } else if (formatStr === "0.0") {
        return { type: "value", value: value.toFixed(1) };
      } else if (formatStr === "0.00") {
        return { type: "value", value: value.toFixed(2) };
      } else if (formatStr === "0.000") {
        return { type: "value", value: value.toFixed(3) };
      } else if (formatStr === "#,##0") {
        return { type: "value", value: Math.round(value).toLocaleString() };
      } else if (formatStr === "#,##0.00") {
        return {
          type: "value",
          value: value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ","),
        };
      } else if (formatStr === "0%") {
        return { type: "value", value: (value * 100).toFixed(0) + "%" };
      } else if (formatStr === "0.0%") {
        return { type: "value", value: (value * 100).toFixed(1) + "%" };
      } else if (formatStr === "0.00%") {
        return { type: "value", value: (value * 100).toFixed(2) + "%" };
      } else {
        // For unrecognized formats, just return the number as string
        return { type: "value", value: value.toString() };
      }
    }

    // For non-numbers, convert to string
    return {
      type: "value",
      value: value === undefined || value === null ? "" : String(value),
    };
  },
};

// Export all text functions
export const textFunctions: FunctionDefinition[] = [
  CONCATENATE,
  LEN,
  UPPER,
  LOWER,
  TRIM,
  LEFT,
  RIGHT,
  MID,
  FIND,
  SEARCH,
  SUBSTITUTE,
  REPLACE,
  FE_CONCAT,
  EXACT,
  TEXT,
];
