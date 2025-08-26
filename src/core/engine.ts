/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  FormulaError,
  type CellAddress,
  type CellNumber,
  type EvaluationContext,
  type SerializedCellValue,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "./types";

import { getCellReference, parseCellReference } from "src/core/utils";
import { Evaluator } from "../evaluator/evaluator";
import { type FunctionEvaluationResult } from "./types";
import { dependencyNodeToKey } from "./utils/dependency-node-key";
import type { FunctionNode } from "src/parser/ast";
import { functions } from "src/functions";

/**
 * Main FormulaEngine class
 */
export class FormulaEngine extends Evaluator {
  constructor() {
    super();
  }

  /**
   * Static factory method to build an empty engine
   */
  static buildEmpty(): FormulaEngine {
    return new FormulaEngine();
  }

  override getCellEvaluationResult(
    cellAddress: CellAddress
  ): FunctionEvaluationResult | undefined {

    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }

    const sheet = this.sheets.get(cellAddress.sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // maybe it is a spilled cell, we need to check the spilled values
    const context: EvaluationContext = {
      currentSheet: cellAddress.sheetName,
      currentCell: cellAddress,
      evaluationStack: new Set(),
      dependencies: new Set(),
    };
    const spilled = this.evaluateSpilled(cellAddress, context);
    if (spilled.isSpilled) {
      return spilled.result;
    }

    this.evaluateCell(cellAddress);

    const value = this.evaluatedNodes.get(
      dependencyNodeToKey({
        type: "cell",
        address: cellAddress,
        sheetName: sheet.name,
      })
    );

    if (!value || !value.evaluationResult) {
      // nothing in the cell
      return undefined;
    }

    return value.evaluationResult;
  }

  getCellValue(cellAddress: CellAddress, debug?: boolean): SerializedCellValue {
    const result = this.getCellEvaluationResult(cellAddress);
    if (!result) {
      return undefined;
    }

    return this.evaluationResultToSerializedValue(result, debug);
  }

  getCellDependents(
    address: CellAddress | SpreadsheetRange
  ): (SpreadsheetRange | CellAddress)[] {
    throw new Error("Not implemented");
  }

  getCellPrecedents(
    address: CellAddress | SpreadsheetRange
  ): (SpreadsheetRange | CellAddress)[] {
    throw new Error("Not implemented");
  }

  addNamedExpression({
    expression,
    expressionName,
    sheetName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
  }) {
    if (!sheetName) {
      this.globalNamedExpressions.set(expressionName, {
        name: expressionName,
        expression,
      });
    } else {
      let scopedNamedExpressions = this.scopedNamedExpressions.get(sheetName);
      if (!scopedNamedExpressions) {
        scopedNamedExpressions = new Map();
        this.scopedNamedExpressions.set(sheetName, scopedNamedExpressions);
      }

      scopedNamedExpressions.set(expressionName, {
        name: expressionName,
        expression,
      });
    }

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  addTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
  }: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    const { rowIndex, colIndex } = parseCellReference(start);
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    const headers = new Map<string, { name: string; index: number }>();
    for (let i = 0; i < numCols; i++) {
      const header = sheet.content.get(
        getCellReference({ rowIndex, colIndex: colIndex + i })
      );

      if (header) {
        headers.set(String(header), { name: String(header), index: i });
      } else {
        headers.set(`Column ${i + 1}`, { name: `Column ${i + 1}`, index: i });
      }
    }

    const endRow: SpreadsheetRangeEnd =
      numRows.type === "number"
        ? { type: "number", value: rowIndex + numRows.value }
        : numRows;

    console.log("EndRow", { endRow, rowIndex, numRows });

    const table: TableDefinition = {
      name: tableName,
      sheetName,
      start: {
        rowIndex,
        colIndex,
      },
      headers,
      endRow,
    };

    let tables = this.tables.get(sheetName);
    if (!tables) {
      tables = new Map();
      this.tables.set(sheetName, tables);
    }

    tables.set(tableName, table);

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return table;
  }

  public setSheetContent(
    sheetName: string,
    content: Map<string, SerializedCellValue>
  ) {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    sheet.content = content;

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  setCellContent(address: CellAddress, content: SerializedCellValue) {
    const sheet = this.sheets.get(address.sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    sheet.content.set(getCellReference(address), content);
    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  triggerCellsUpdateEvent() {
    this.cellsUpdateListeners.forEach((sheetListeners) =>
      sheetListeners.forEach((listener) => listener())
    );
  }

  reevaluateSheet(sheetName: string) {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    for (const key of sheet.content.keys()) {
      const address = parseCellReference(key);
      this.evaluateCell({ ...address, sheetName });
    }
  }

  /**
   * Re-evaluates all sheets to ensure all dependencies are resolved correctly
   */
  reevaluate() {
    this.evaluatedNodes.clear();
    this.spilledValues.clear();
    for (const sheet of this.sheets.values()) {
      this.reevaluateSheet(sheet.name);
    }
  }

  override evaluateFunction(
    node: FunctionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    const func = functions[node.name];
    if (!func) {
      throw new Error(FormulaError.NAME);
    }
    return func.evaluate.call(this, node, context);
  }
}
