/**
 * Sheet renamer utility for updating cross-sheet references in formulas
 */

import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import { transformAST } from "./ast-traverser";
import type { ASTNode } from "../parser/ast";

/**
 * Renames sheet references in a formula string
 * @param formula The formula string (without the leading =)
 * @param oldSheetName The old sheet name to replace
 * @param newSheetName The new sheet name to use
 * @returns The updated formula string
 */
export function renameSheetInFormula(options: {
  formula: string,
  oldSheetName: string;
  newSheetName: string;
}): string {
  const { formula, oldSheetName, newSheetName } = options;
  try {
    const ast = parseFormula(formula);
    
    const updatedAst = transformAST(ast, (node) => {
      // Handle regular cross-sheet references (e.g., Sheet1!A1)
      if (node.type === "reference" && node.sheetName === oldSheetName) {
        return {
          ...node,
          sheetName: newSheetName,
        };
      }
      
      // Handle range references with sheet names (e.g., Sheet1!A1:B2)
      if (node.type === "range" && node.sheetName === oldSheetName) {
        return {
          ...node,
          sheetName: newSheetName,
        };
      }
      
      return node;
    });
    
    return astToString(updatedAst);
  } catch (error) {
    // If parsing fails, return the original formula
    return formula;
  }
}

/**
 * Checks if a formula references a specific sheet
 * @param formula The formula string (without the leading =)
 * @param sheetName The sheet name to check for
 * @returns True if the formula references the sheet
 */
export function formulaReferencesSheet(formula: string, sheetName: string): boolean {
  try {
    const ast = parseFormula(formula);
    const referencedSheets = getReferencedSheetNames(formula);
    return referencedSheets.includes(sheetName);
  } catch (error) {
    // If parsing fails, assume no reference
    return false;
  }
}

/**
 * Gets all sheet names referenced in a formula
 * @param formula The formula string (without the leading =)
 * @returns Array of unique sheet names referenced in the formula
 */
export function getReferencedSheetNames(formula: string): string[] {
  try {
    const ast = parseFormula(formula);
    const sheetNames = new Set<string>();

    transformAST(ast, (node) => {
      // Handle cross-sheet references
      if ((node.type === "reference" || node.type === "range") && node.sheetName) {
        sheetNames.add(node.sheetName);
      }
      
      return node;
    });

    return Array.from(sheetNames);
  } catch (error) {
    // If parsing fails, return empty array
    return [];
  }
}

