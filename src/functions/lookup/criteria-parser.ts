import { parseFormula } from "src/parser/parser";
import type { ASTNode } from "src/parser/ast";
import { FormulaError, type CellValue } from "src/core/types";

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

    // Try to parse the value part as a formula to get the proper type
    try {
      const ast = parseFormula(valueStr.trim());
      if (ast.type === "value") {
        return { 
          type: "comparison", 
          operator: operator as ">" | "<" | ">=" | "<=" | "<>", 
          value: ast.value 
        };
      }
    } catch {
      // If parsing fails, treat as string
    }

    // Fallback: treat as string value
    return { 
      type: "comparison", 
      operator: operator as ">" | "<" | ">=" | "<=" | "<>", 
      value: { type: "string", value: valueStr } 
    };
  }

  // Check for wildcards
  if (criteriaStr.includes("*") || criteriaStr.includes("?")) {
    return { type: "wildcard", pattern: criteriaStr };
  }

  // Try to parse as a formula to get the proper type
  try {
    const ast = parseFormula(criteriaStr);
    if (ast.type === "value") {
      return { type: "exact", value: ast.value };
    }
  } catch {
    // If parsing fails, treat as string
  }

  // Default: exact string match
  return { type: "exact", value: criteria };
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
      if (parsedCriteria.value.type !== "number") {
        return false; // Can only compare with numeric criteria
      }

      const criteriaNum = parsedCriteria.value.value;

      // Handle infinity cases
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
