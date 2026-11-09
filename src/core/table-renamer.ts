import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import type { StructuredReferenceNode } from "../parser/ast";
import { transformAST } from "./ast-traverser";

/**
 * Whether a structured reference that names a table refers to the table
 * `tableName` of workbook `tableWorkbookName`. Tables are workbook resources:
 * `[Book]Sales[Qty]` names the table in Book from any workbook, and an
 * unqualified `Sales[Qty]` names the table in the formula's own workbook.
 * Without `formulaWorkbookName` (a global named expression is evaluated in the
 * workbook of the cell that uses it) an unqualified reference can refer to any
 * workbook's table, so it matches.
 */
function namesWorkbookTable(
  node: StructuredReferenceNode,
  tableName: string,
  tableWorkbookName: string,
  formulaWorkbookName: string | undefined
): boolean {
  if (node.tableName !== tableName) {
    return false;
  }
  const referencedWorkbookName = node.workbookName ?? formulaWorkbookName;
  return (
    referencedWorkbookName === undefined ||
    referencedWorkbookName === tableWorkbookName
  );
}

/**
 * Renames references to one workbook-scoped table in a formula string.
 * References to a table of the same name in another workbook are unchanged.
 * @returns The updated formula string, or the original if it cannot be parsed
 */
export function renameTableInFormula(options: {
  /** The formula string (without the leading =) */
  formula: string;
  oldTableName: string;
  newTableName: string;
  /** The workbook that owns the renamed table. */
  tableWorkbookName: string;
  /**
   * The workbook the formula belongs to, which unqualified references resolve
   * in. Omit it for a formula that belongs to no workbook, such as a global
   * named expression.
   */
  formulaWorkbookName?: string;
}): string {
  const {
    formula,
    oldTableName,
    newTableName,
    tableWorkbookName,
    formulaWorkbookName,
  } = options;
  try {
    const ast = parseFormula(formula);
    
    const transformedAST = transformAST(ast, (node) => {
      if (
        node.type === "structured-reference" &&
        namesWorkbookTable(
          node,
          oldTableName,
          tableWorkbookName,
          formulaWorkbookName
        )
      ) {
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
        ? namesWorkbookTable(
            node,
            tableName,
            tableWorkbookName,
            formulaWorkbookName
          )
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
