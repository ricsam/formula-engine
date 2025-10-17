import type { CellAddress } from "src/core/types";
import type { CellEvalNode } from "./cell-eval-node";
import type { RangeEvaluationNode } from "./range-evaluation-node";
import type { EmptyCellEvaluationNode } from "./empty-cell-evaluation-node";
import type { DependencyNode } from "src/core/managers/dependency-node";

export class EvaluationContext {
  /**
   * Can be a range or a cell
   */
  private _dependencyNode: DependencyNode;
  /**
   * The cell evaluating a cell,e.g.
   * if we are evaluting A1=SUM(B2:B4) + B1, then the origin cell is A1 and the dependency node is A1 as well
   * the open range evaluator will create a new context with the origin cell being A1 and the dependency node being B2:B4
   *
   * A new dependency will be added to A1 onto B1, and then B1 will be evaluated just like A1 is evaluated where the origin cell is B1
   */
  private _originCell: CellEvalNode | EmptyCellEvaluationNode;
  private _tableName?: string;
  constructor(
    dependencyNode: DependencyNode,
    originCell: CellEvalNode | EmptyCellEvaluationNode,
    tableName?: string
  ) {
    this._dependencyNode = dependencyNode;
    this._originCell = originCell;
    this._tableName = tableName;
  }

  get dependencyNode() {
    return this._dependencyNode;
  }

  get originCell() {
    return this._originCell;
  }

  private _contextDependency: ContextDependency = {};

  getContextDependency() {
    return this._contextDependency;
  }

  addContextDependency(...types: ContextDependencyType[]) {
    types.forEach((type) => {
      switch (type) {
        case "row":
          this._contextDependency.rowIndex =
            this.originCell.cellAddress.rowIndex;
          break;
        case "col":
          this._contextDependency.colIndex =
            this.originCell.cellAddress.colIndex;
          break;
        case "workbook":
          this._contextDependency.workbookName =
            this.originCell.cellAddress.workbookName;
          break;
        case "sheet":
          this._contextDependency.sheetName =
            this.originCell.cellAddress.sheetName;
          break;
        case "table":
          this._contextDependency.tableName = this._tableName;
          break;
      }
    });
  }

  /**
   * When evaluating an AST node,
   * we need to append the subtree context
   * dependencies to the current context dependency
   */
  appendContextDependency(contextDependency: ContextDependency) {
    this._contextDependency = {
      ...this._contextDependency,
      ...Object.fromEntries(
        Object.entries(contextDependency).filter(
          ([key, value]) => value !== undefined
        )
      ),
    };
  }
}

/**
 * Each value has the same value as the origin cell
 * the defined keys are the ones the ast node is dependent on
 * e.g. A3=ROW() will have a context dependency of { rowIndex: 3 }
 *
 * The keys are ANDed together, e.g. { workbookName: "Sheet1", sheetName: "Sheet2" }
 * means the ast node is dependent on the workbook "Sheet1" and the sheet "Sheet2"
 */
export type ContextDependency = {
  workbookName?: string;
  sheetName?: string;
  tableName?: string;
  rowIndex?: number;
  colIndex?: number;
};

export const contextDependencyKeys = [
  "workbookName",
  "sheetName",
  "tableName",
  "rowIndex",
  "colIndex",
] as const;

/**
 * These are some distinct scenarios where context dependencies are added
 */
export type ContextDependencyType =
  | "row"
  | "col"
  | "workbook"
  | "sheet"
  | "table";

// *  [astKey], // `=1+1`
// *  [astKey, sheetKey, workbookKey], // `B3`
// *  [astKey, workbookKey], // `Table1[Column1]`
// *  [astKey, workbookKey], // `Sheet1!B3`
// *  [astKey, cellAddress.rowIndex], // `ROW()`
// *  [astKey, cellAddress.colIndex], // `COL()`
// *  [astKey, cellAddress.rowIndex, cellAddress.colIndex] // `CELL("address")`
// *  [astKey, tableKey, cellAddress.rowIndex], // `@Column1`
// *  [astKey, workbookKey, cellAddress.rowIndex], // `Table1[@Column1]`
