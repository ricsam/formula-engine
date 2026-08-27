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
 * Renames structured references to a column in one workbook-scoped table.
 * Implicit references (for example [@Price]) are only changed when the caller
 * confirms that the formula belongs to the table being updated.
 */
export function renameTableColumnInFormula(options: {
  formula: string;
  tableName: string;
  tableWorkbookName: string;
  formulaWorkbookName?: string;
  oldColumnName: string;
  newColumnName: string;
  includeImplicitReferences?: boolean;
}): string {
  return renameTableColumnsInFormula({
    formula: options.formula,
    tableName: options.tableName,
    tableWorkbookName: options.tableWorkbookName,
    formulaWorkbookName: options.formulaWorkbookName,
    columnRenames: new Map([
      [options.oldColumnName, options.newColumnName],
    ]),
    includeImplicitReferences: options.includeImplicitReferences,
  });
}

export function renameTableColumnsInFormula(options: {
  formula: string;
  tableName: string;
  tableWorkbookName: string;
  formulaWorkbookName?: string;
  columnRenames: ReadonlyMap<string, string>;
  includeImplicitReferences?: boolean;
}): string {
  const {
    formula,
    tableName,
    tableWorkbookName,
    formulaWorkbookName,
    columnRenames,
    includeImplicitReferences = false,
  } = options;

  if (
    columnRenames.size === 0 ||
    Array.from(columnRenames).every(([oldName, newName]) => oldName === newName)
  ) {
    return formula;
  }

  try {
    const ast = parseFormula(formula);
    let changed = false;
    const transformedAST = transformAST(ast, (node) => {
      if (node.type !== "structured-reference" || !node.cols) {
        return node;
      }

      const referencesTargetTable = node.tableName
        ? node.tableName === tableName &&
          (node.workbookName
            ? node.workbookName === tableWorkbookName
            : formulaWorkbookName === undefined ||
              formulaWorkbookName === tableWorkbookName)
        : includeImplicitReferences;
      if (!referencesTargetTable) {
        return node;
      }

      const startCol = columnRenames.get(node.cols.startCol) ?? node.cols.startCol;
      const endCol = columnRenames.get(node.cols.endCol) ?? node.cols.endCol;
      if (startCol === node.cols.startCol && endCol === node.cols.endCol) {
        return node;
      }

      changed = true;
      return {
        ...node,
        cols: { startCol, endCol },
      };
    });

    return changed ? astToString(transformedAST) : formula;
  } catch {
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
