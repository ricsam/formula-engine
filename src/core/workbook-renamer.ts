/**
 * Workbook renamer utility for updating cross-workbook references in formulas
 */

import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import { transformAST } from "./ast-traverser";
import type { ASTNode } from "../parser/ast";

/**
 * Renames workbook references in a formula string
 * @param formula The formula string (without the leading =)
 * @param oldWorkbookName The old workbook name to replace
 * @param newWorkbookName The new workbook name to use
 * @returns The updated formula string
 */
export function renameWorkbookInFormula(options: {
  formula: string;
  oldWorkbookName: string;
  newWorkbookName: string;
}): string {
  const { formula, oldWorkbookName, newWorkbookName } = options;
  try {
    const ast = parseFormula(formula);
    
    const updatedAst = transformAST(ast, (node) => {
      // Handle cross-workbook references (e.g., [MyWorkbook]Sheet1!A1)
      if (node.type === "reference" && node.workbookName === oldWorkbookName) {
        return {
          ...node,
          workbookName: newWorkbookName,
        };
      }
      
      // Handle range references with workbook names (e.g., [MyWorkbook]Sheet1!A1:B2)
      if (node.type === "range" && node.workbookName === oldWorkbookName) {
        return {
          ...node,
          workbookName: newWorkbookName,
        };
      }
      
      // Handle named expressions with workbook names
      if (node.type === "named-expression" && node.workbookName === oldWorkbookName) {
        return {
          ...node,
          workbookName: newWorkbookName,
        };
      }
      
      // Handle structured references with workbook names (e.g., [MyWorkbook]Sheet1!Table1[Column1])
      if (node.type === "structured-reference" && node.workbookName === oldWorkbookName) {
        return {
          ...node,
          workbookName: newWorkbookName,
        };
      }
      
      // Handle 3D ranges with workbook names (e.g., [MyWorkbook]Sheet1:Sheet3!A1:C5)
      if (node.type === "3d-range" && node.workbookName === oldWorkbookName) {
        return {
          ...node,
          workbookName: newWorkbookName,
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
 * Checks if a formula references a specific workbook
 * @param formula The formula string (without the leading =)
 * @param workbookName The workbook name to check for
 * @returns True if the formula references the workbook
 */
export function formulaReferencesWorkbook(formula: string, workbookName: string): boolean {
  try {
    const referencedWorkbooks = getReferencedWorkbookNames(formula);
    return referencedWorkbooks.includes(workbookName);
  } catch (error) {
    // If parsing fails, assume no reference
    return false;
  }
}

/**
 * Gets all workbook names referenced in a formula
 * @param formula The formula string (without the leading =)
 * @returns Array of unique workbook names referenced in the formula
 */
export function getReferencedWorkbookNames(formula: string): string[] {
  try {
    const ast = parseFormula(formula);
    const workbookNames = new Set<string>();

    transformAST(ast, (node) => {
      // Handle cross-workbook references
      if ((node.type === "reference" || 
           node.type === "range" || 
           node.type === "named-expression" || 
           node.type === "structured-reference" ||
           node.type === "3d-range") && 
          node.workbookName) {
        workbookNames.add(node.workbookName);
      }
      
      return node;
    });

    return Array.from(workbookNames);
  } catch (error) {
    // If parsing fails, return empty array
    return [];
  }
}
