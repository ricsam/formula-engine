import { FormulaError, type CellValue } from "../core/types";

/**
 * Parsed criteria result
 */
export type ParsedCriteria = 
  | { type: "exact"; value: CellValue }
  | { type: "comparison"; operator: ">" | "<" | ">=" | "<=" | "<>"; value: CellValue }
  | { type: "wildcard"; pattern: string }
  | { type: "error"; message: string };

/**
 * Convert wildcard pattern to regex
 */
function wildcardToRegex(pattern: string): RegExp {
  // Escape regex special characters except * and ?
  const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&");
  // Convert wildcards: * -> .*, ? -> .
  const regexPattern = escaped.replace(/\*/g, ".*").replace(/\?/g, ".");
  return new RegExp(`^${regexPattern}$`);
}

/**
 * Parse a simple value string to determine its type
 * This is a lightweight alternative to parseFormula for criteria strings
 */
function parseSimpleValue(valueStr: string): CellValue {
  // Trim the value
  const trimmed = valueStr.trim();
  
  // Check for boolean values (case-insensitive)
  const upperTrimmed = trimmed.toUpperCase();
  if (upperTrimmed === "TRUE") {
    return { type: "boolean", value: true };
  }
  if (upperTrimmed === "FALSE") {
    return { type: "boolean", value: false };
  }
  
  // Check for infinity values (case-insensitive)
  if (upperTrimmed === "INFINITY") {
    return { type: "infinity", sign: "positive" };
  }
  if (upperTrimmed === "-INFINITY") {
    return { type: "infinity", sign: "negative" };
  }
  
  // Check for numeric values
  // Match: optional negative sign, digits, optional decimal point and more digits
  // Examples: "10", "-20", "3.14", "-0.5", "0"
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (!isNaN(num) && isFinite(num)) {
      return { type: "number", value: num };
    }
  }
  
  // Default to string
  return { type: "string", value: valueStr };
}

/**
 * Simple criteria parser for COUNTIF
 * Handles: exact values, comparisons (>5, <=10, <>0), and wildcards (*app, ?at)
 */
export function parseCriteria(criteria: CellValue): ParsedCriteria {
  // Only string criteria can contain operators or wildcards
  if (criteria.type !== "string") {
    return { type: "exact", value: criteria };
  }

  const criteriaStr = criteria.value;

  // Handle empty string as exact match
  if (criteriaStr === "") {
    return { type: "exact", value: criteria };
  }

  // Special case: "=" means count empty cells
  if (criteriaStr === "=") {
    return { type: "exact", value: { type: "string", value: "" } };
  }

  // Check for comparison operators (must be at the beginning)
  const comparisonMatch = criteriaStr.match(/^(>=|<=|<>|>|<)(.*)$/);
  if (comparisonMatch) {
    const operator = comparisonMatch[1];
    const valueStr = comparisonMatch[2];
    
    if (!operator || valueStr === undefined) {
      return { type: "error", message: "Invalid comparison operator format" };
    }
    
    // Special case: "<>" with no value means "not empty"
    if (operator === "<>" && valueStr === "") {
      return { 
        type: "comparison", 
        operator: "<>", 
        value: { type: "string", value: "" } 
      };
    }

    // Parse the value part to get the proper type
    const parsedValue = parseSimpleValue(valueStr);
    return { 
      type: "comparison", 
      operator: operator as ">" | "<" | ">=" | "<=" | "<>", 
      value: parsedValue
    };
  }

  // Check for wildcards
  if (criteriaStr.includes("*") || criteriaStr.includes("?")) {
    return { type: "wildcard", pattern: criteriaStr };
  }

  // Parse as a simple value to get the proper type
  const parsedValue = parseSimpleValue(criteriaStr);
  return { type: "exact", value: parsedValue };
}

/**
 * Check if a cell value matches the parsed criteria
 */
export function matchesParsedCriteria(cellValue: CellValue, parsedCriteria: ParsedCriteria): boolean {
  switch (parsedCriteria.type) {
    case "error":
      return false;

    case "exact":
      // Exact match - must be same type and value
      if (cellValue.type !== parsedCriteria.value.type) {
        return false;
      }
      if ("value" in cellValue && "value" in parsedCriteria.value) {
        if (cellValue.type === 'string' && parsedCriteria.value.type === 'string') {
          return cellValue.value.toLowerCase() === parsedCriteria.value.value.toLowerCase();
        }
        return cellValue.value === parsedCriteria.value.value;
      }
      if (cellValue.type === "infinity" && parsedCriteria.value.type === "infinity") {
        return cellValue.sign === parsedCriteria.value.sign;
      }
      return false;

    case "comparison":
      // Special case: "<>" with empty string means "not empty"
      if (parsedCriteria.operator === "<>" && 
          parsedCriteria.value.type === "string" && 
          parsedCriteria.value.value === "") {
        return cellValue.type !== "string" || cellValue.value !== "";
      }

      // For other comparisons, handle by operator
      if (parsedCriteria.operator === "<>") {
        // Not equals - same logic as exact but negated
        if (cellValue.type !== parsedCriteria.value.type) {
          return true; // Different types are not equal
        }
        if ("value" in cellValue && "value" in parsedCriteria.value) {
          return cellValue.value !== parsedCriteria.value.value;
        }
        if (cellValue.type === "infinity" && parsedCriteria.value.type === "infinity") {
          return cellValue.sign !== parsedCriteria.value.sign;
        }
        return true;
      }

      // Numeric comparison operators - only work with numbers and infinity
      if (cellValue.type !== "number" && cellValue.type !== "infinity") {
        return false;
      }
      if (parsedCriteria.value.type !== "number" && parsedCriteria.value.type !== "infinity") {
        return false; // Can only compare with numeric or infinity criteria
      }

      // Handle when criteria is infinity
      if (parsedCriteria.value.type === "infinity") {
        const criteriaSign = parsedCriteria.value.sign;
        
        // Handle when cell value is also infinity
        if (cellValue.type === "infinity") {
          const cellSign = cellValue.sign;
          switch (parsedCriteria.operator) {
            case ">":
              // +∞ > -∞, but not +∞ > +∞
              return cellSign === "positive" && criteriaSign === "negative";
            case "<":
              // -∞ < +∞, but not -∞ < -∞
              return cellSign === "negative" && criteriaSign === "positive";
            case ">=":
              // +∞ >= +∞, +∞ >= -∞, but not -∞ >= +∞
              return cellSign === "positive" || (cellSign === criteriaSign);
            case "<=":
              // -∞ <= -∞, -∞ <= +∞, but not +∞ <= -∞
              return cellSign === "negative" || (cellSign === criteriaSign);
            default:
              return false;
          }
        }
        
        // Handle when cell value is a finite number
        if (cellValue.type === "number") {
          switch (parsedCriteria.operator) {
            case ">":
              // No finite number is > +∞, all finite numbers > -∞
              return criteriaSign === "negative";
            case "<":
              // All finite numbers < +∞, no finite number < -∞
              return criteriaSign === "positive";
            case ">=":
              // No finite number >= +∞, all finite numbers >= -∞
              return criteriaSign === "negative";
            case "<=":
              // All finite numbers <= +∞, no finite number <= -∞
              return criteriaSign === "positive";
            default:
              return false;
          }
        }
      }

      // Handle when criteria is a number (not infinity)
      if (parsedCriteria.value.type === "number") {
        const criteriaNum = parsedCriteria.value.value;

        // Handle infinity cell values with numeric criteria
        if (cellValue.type === "infinity") {
          switch (parsedCriteria.operator) {
            case ">":
              return cellValue.sign === "positive"; // +∞ > any number
            case "<":
              return cellValue.sign === "negative"; // -∞ < any number
            case ">=":
              return cellValue.sign === "positive"; // +∞ >= any number
            case "<=":
              return cellValue.sign === "negative"; // -∞ <= any number
            default:
              return false;
          }
        }

        // Regular number comparison
        if (cellValue.type === "number") {
          switch (parsedCriteria.operator) {
            case ">":
              return cellValue.value > criteriaNum;
            case "<":
              return cellValue.value < criteriaNum;
            case ">=":
              return cellValue.value >= criteriaNum;
            case "<=":
              return cellValue.value <= criteriaNum;
            default:
              return false;
          }
        }
      }
      
      return false;

    case "wildcard":
      // Wildcard matching - only works with strings
      if (cellValue.type !== "string") {
        return false;
      }
      const regex = wildcardToRegex(parsedCriteria.pattern);
      return regex.test(cellValue.value);

    default:
      return false;
  }
}
