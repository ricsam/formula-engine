import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import { transformAST } from "./ast-traverser";

/**
 * Renames table references in a formula string
 * @param formula - The formula string (without the leading =)
 * @param oldTableName - The current table name to replace
 * @param newTableName - The new table name
 * @returns The updated formula string, or the original if no changes were made
 */
export function renameTableInFormula(
  formula: string,
  oldTableName: string,
  newTableName: string
): string {
  try {
    const ast = parseFormula(formula);
    
    const transformedAST = transformAST(ast, (node) => {
      if (node.type === "structured-reference" && node.tableName === oldTableName) {
        return {
          ...node,
          tableName: newTableName
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
 * Checks if a formula contains references to a specific table
 * @param formula - The formula string (without the leading =)
 * @param tableName - The table name to search for
 * @returns True if the formula contains references to the table
 */
export function formulaReferencesTable(formula: string, tableName: string): boolean {
  try {
    const ast = parseFormula(formula);
    let hasReference = false;

    transformAST(ast, (node) => {
      if (node.type === "structured-reference" && node.tableName === tableName) {
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
 * Gets all table names referenced in a formula
 * @param formula - The formula string (without the leading =)
 * @returns Array of unique table names referenced in the formula
 */
export function getReferencedTableNames(formula: string): string[] {
  try {
    const ast = parseFormula(formula);
    const tableNames = new Set<string>();

    transformAST(ast, (node) => {
      if (node.type === "structured-reference" && node.tableName) {
        tableNames.add(node.tableName);
      }
      return node;
    });

    return Array.from(tableNames);
  } catch (error) {
    // If parsing fails, return empty array
    return [];
  }
}
