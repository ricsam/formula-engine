import {
  createErrorNode,
  type ASTNode,
  type RangeNode,
  type ReferenceNode,
} from "../parser/ast";
import { astToString } from "../parser/formatter";
import { parseFormula } from "../parser/parser";
import { renderCompactFormulaBody } from "../language/formula-formatter";
import { transformAST, traverseAST } from "./ast-traverser";
import { FormulaError } from "./types";

export interface ShiftFormulaReferencesOptions {
  /**
   * Keep ranges that span whole columns (`A:A`, which the engine stores as the
   * open-ended range `A1:A`) or whole rows (`1:1`, stored as `A1:1`) anchored at
   * the sheet edge instead of moving that edge. Sorting a table uses this: a
   * row moves with its formulas, but `COUNTIF(A:A, [@Name])` must still count
   * the whole column.
   */
  preserveWholeRowsAndColumns?: boolean;
  /**
   * Return the formula in the compact form the engine stores cell formulas in,
   * as `normalizeFormulaCellContent` would, without parsing it a second time.
   */
  normalize?: boolean;
}

/**
 * Whether a formula contains A1 references or ranges, the only parts that
 * {@link shiftFormulaReferences} can change. Unparseable formulas have none.
 */
export function hasPositionalReferences(formula: string): boolean {
  if (!formula.startsWith("=")) {
    return false;
  }
  let found = false;
  try {
    traverseAST(parseFormula(formula.slice(1)), (node) => {
      if (node.type === "reference" || node.type === "range") {
        found = true;
      }
    });
  } catch {
    return false;
  }
  return found;
}

/**
 * Moves the relative references of a formula by a row and column offset, the
 * way copying the formula to another cell does. Absolute coordinates stay put
 * and a reference that would leave the sheet becomes `#REF!`, as in Excel.
 * Structured references and names are not positional and stay unchanged.
 *
 * Returns the formula unchanged when the offset is zero or it cannot be parsed.
 */
export function shiftFormulaReferences(
  formula: string,
  rowDelta: number,
  colDelta: number,
  options: ShiftFormulaReferencesOptions = {}
): string {
  if (!formula.startsWith("=") || (rowDelta === 0 && colDelta === 0)) {
    return formula;
  }

  let ast: ASTNode;
  try {
    ast = parseFormula(formula.slice(1));
  } catch {
    return formula;
  }

  const offSheet = () =>
    createErrorNode(FormulaError.REF, "Reference moved off the sheet");

  const shiftReference = (node: ReferenceNode): ASTNode => {
    const colIndex = node.isAbsolute.col
      ? node.address.colIndex
      : node.address.colIndex + colDelta;
    const rowIndex = node.isAbsolute.row
      ? node.address.rowIndex
      : node.address.rowIndex + rowDelta;
    if (colIndex < 0 || rowIndex < 0) {
      return offSheet();
    }
    return { ...node, address: { colIndex, rowIndex } };
  };

  const shiftRange = (node: RangeNode): ASTNode => {
    const { range, isAbsolute } = node;
    const keepFirstRow =
      options.preserveWholeRowsAndColumns === true &&
      range.end.row.type === "infinity" &&
      range.start.row === 0;
    const keepFirstCol =
      options.preserveWholeRowsAndColumns === true &&
      range.end.col.type === "infinity" &&
      range.start.col === 0;
    const startRow =
      isAbsolute.start.row || keepFirstRow
        ? range.start.row
        : range.start.row + rowDelta;
    const startCol =
      isAbsolute.start.col || keepFirstCol
        ? range.start.col
        : range.start.col + colDelta;
    const endRow =
      range.end.row.type === "number" && !isAbsolute.end.row
        ? { type: "number" as const, value: range.end.row.value + rowDelta }
        : range.end.row;
    const endCol =
      range.end.col.type === "number" && !isAbsolute.end.col
        ? { type: "number" as const, value: range.end.col.value + colDelta }
        : range.end.col;
    if (
      startRow < 0 ||
      startCol < 0 ||
      (endRow.type === "number" && endRow.value < 0) ||
      (endCol.type === "number" && endCol.value < 0)
    ) {
      return offSheet();
    }
    return {
      ...node,
      range: {
        start: { row: startRow, col: startCol },
        end: { row: endRow, col: endCol },
      },
    };
  };

  const shifted = transformAST(ast, (node) => {
    switch (node.type) {
      case "reference":
        return shiftReference(node);
      case "range":
        return shiftRange(node);
      case "3d-range":
        // A sheet span whose cell part left the sheet is #REF! as a whole.
        return node.reference.type === ("error" as string) ? offSheet() : node;
      default:
        return node;
    }
  });

  return `=${
    options.normalize ? renderCompactFormulaBody(shifted) : astToString(shifted)
  }`;
}
