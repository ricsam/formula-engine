import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import { transformAST } from "./ast-traverser";

/**
 * Renames named expression references in a formula string
 * @param formula - The formula string (without the leading =)
 * @param oldName - The current named expression name to replace
 * @param newName - The new named expression name
 * @returns The updated formula string, or the original if no changes were made
 */
export function renameNamedExpressionInFormula(
  formula: string,
  oldName: string,
  newName: string
): string {
  try {
    const ast = parseFormula(formula);
    
    const transformedAST = transformAST(ast, (node) => {
      if (node.type === "named-expression" && node.name === oldName) {
        return {
          ...node,
          name: newName
        };
      }
      return node;
    });

    return astToString(transformedAST);
  } catch (error) {
    // If parsing fails, return the original formula
    return formula;
  }
}

/**
 * Checks if a formula contains references to a specific named expression
 * @param formula - The formula string (without the leading =)
 * @param namedExpressionName - The named expression name to search for
 * @returns True if the formula contains references to the named expression
 */
export function formulaReferencesNamedExpression(formula: string, namedExpressionName: string): boolean {
  try {
    const ast = parseFormula(formula);
    let hasReference = false;

    transformAST(ast, (node) => {
      if (node.type === "named-expression" && node.name === namedExpressionName) {
        hasReference = true;
      }
      return node;
    });

    return hasReference;
  } catch (error) {
    // If parsing fails, assume no reference
    return false;
  }
}

/**
 * Gets all named expression names referenced in a formula
 * @param formula - The formula string (without the leading =)
 * @returns Array of unique named expression names referenced in the formula
 */
export function getReferencedNamedExpressionNames(formula: string): string[] {
  try {
    const ast = parseFormula(formula);
    const namedExpressionNames = new Set<string>();

    transformAST(ast, (node) => {
      if (node.type === "named-expression" && node.name) {
        namedExpressionNames.add(node.name);
      }
      return node;
    });

    return Array.from(namedExpressionNames);
  } catch (error) {
    // If parsing fails, return empty array
    return [];
  }
}
