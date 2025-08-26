/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  FormulaError,
  type CellAddress,
  type CellNumber,
  type EvaluationContext,
  type NamedExpression,
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
import { parseFormula } from "src/parser/parser";
import { astToString, formatFormula } from "src/parser/formatter";
import { renameTableInFormula } from "./table-renamer";
import { renameSheetInFormula } from "./sheet-renamer";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";

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
      this.emit(
        "global-named-expressions-updated",
        this.globalNamedExpressions
      );
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

  removeNamedExpression({
    expressionName,
    sheetName,
  }: {
    expressionName: string;
    sheetName?: string;
  }) {
    let found = false;

    if (!sheetName) {
      // Remove from global named expressions
      found = this.globalNamedExpressions.delete(expressionName);
      if (found) {
        this.emit(
          "global-named-expressions-updated",
          this.globalNamedExpressions
        );
      }
    } else {
      // Remove from sheet-scoped named expressions
      const scopedNamedExpressions = this.scopedNamedExpressions.get(sheetName);
      if (scopedNamedExpressions) {
        found = scopedNamedExpressions.delete(expressionName);
      }
    }

    if (found) {
      // Re-evaluate all sheets since named expressions can be referenced from anywhere
      this.reevaluate();
      this.triggerCellsUpdateEvent();
    }

    return found;
  }

  updateNamedExpression({
    expression,
    expressionName,
    sheetName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
  }) {
    // Check if the named expression exists
    const exists = sheetName
      ? this.scopedNamedExpressions.get(sheetName)?.has(expressionName)
      : this.globalNamedExpressions.has(expressionName);

    if (!exists) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Update is the same as add for existing expressions
    this.addNamedExpression({ expression, expressionName, sheetName });
  }

  renameNamedExpression({
    expressionName,
    sheetName,
    newName,
  }: {
    expressionName: string;
    sheetName?: string;
    newName: string;
  }) {
    // Check if the named expression exists
    const targetMap = sheetName
      ? this.scopedNamedExpressions.get(sheetName)
      : this.globalNamedExpressions;

    if (!targetMap || !targetMap.has(expressionName)) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Check if the new name already exists
    if (targetMap.has(newName)) {
      throw new Error(`Named expression '${newName}' already exists`);
    }

    // Get the expression to rename
    const namedExpression = targetMap.get(expressionName)!;

    // Update the name and re-add with new name
    const updatedExpression = { ...namedExpression, name: newName };
    targetMap.set(newName, updatedExpression);
    targetMap.delete(expressionName);

    // Update all formulas that reference this named expression in sheet cells
    this.sheets.forEach((sheet) => {
      sheet.content.forEach((cell, key) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          const formula = cell.slice(1);
          const updatedFormula = renameNamedExpressionInFormula(
            formula,
            expressionName,
            newName
          );

          // Only update if the formula actually changed
          if (updatedFormula !== formula) {
            sheet.content.set(key, `=${updatedFormula}`);
          }
        }
      });
    });

    // Update global named expressions that reference this named expression
    this.globalNamedExpressions.forEach((namedExpr, name) => {
      if (name !== expressionName) {
        // Don't update the expression we're renaming
        const updatedExpression = renameNamedExpressionInFormula(
          namedExpr.expression,
          expressionName,
          newName
        );

        if (updatedExpression !== namedExpr.expression) {
          this.globalNamedExpressions.set(name, {
            ...namedExpr,
            expression: updatedExpression,
          });
        }
      }
    });

    // Update scoped named expressions that reference this named expression
    this.scopedNamedExpressions.forEach((namedExpressionsMap, sheetName) => {
      namedExpressionsMap.forEach((namedExpr, name) => {
        if (name !== expressionName) {
          // Don't update the expression we're renaming
          const updatedExpression = renameNamedExpressionInFormula(
            namedExpr.expression,
            expressionName,
            newName
          );

          if (updatedExpression !== namedExpr.expression) {
            namedExpressionsMap.set(name, {
              ...namedExpr,
              expression: updatedExpression,
            });
          }
        }
      });
    });

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.triggerCellsUpdateEvent();
    this.emit("global-named-expressions-updated", this.globalNamedExpressions);

    return true;
  }

  makeTable({
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

    return table;
  }

  addTable(props: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    const tableName = props.tableName;
    const table = this.makeTable(props);

    this.tables.set(tableName, table);

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();
    this.emit("tables-updated", this.tables);

    return table;
  }

  renameTable(names: { oldName: string; newName: string }) {
    const table = this.tables.get(names.oldName);
    if (!table) {
      throw new Error("Table not found");
    }
    table.name = names.newName;
    this.tables.set(names.newName, table);
    this.tables.delete(names.oldName);

    // Update all formulas that reference this table in sheet cells
    this.sheets.forEach((sheet) => {
      sheet.content.forEach((cell, key) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          const formula = cell.slice(1);
          const updatedFormula = renameTableInFormula(
            formula,
            names.oldName,
            names.newName
          );

          // Only update if the formula actually changed
          if (updatedFormula !== formula) {
            sheet.content.set(key, `=${updatedFormula}`);
          }
        }
      });
    });

    // Update global named expressions that reference this table
    this.globalNamedExpressions.forEach((namedExpr, name) => {
      const updatedExpression = renameTableInFormula(
        namedExpr.expression,
        names.oldName,
        names.newName
      );

      // Only update if the expression actually changed
      if (updatedExpression !== namedExpr.expression) {
        this.globalNamedExpressions.set(name, {
          ...namedExpr,
          expression: updatedExpression,
        });
      }
    });

    // Update sheet-scoped named expressions that reference this table
    this.scopedNamedExpressions.forEach((namedExpressionsMap, sheetName) => {
      namedExpressionsMap.forEach((namedExpr, name) => {
        const updatedExpression = renameTableInFormula(
          namedExpr.expression,
          names.oldName,
          names.newName
        );

        // Only update if the expression actually changed
        if (updatedExpression !== namedExpr.expression) {
          namedExpressionsMap.set(name, {
            ...namedExpr,
            expression: updatedExpression,
          });
        }
      });
    });

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();
    this.emit("tables-updated", this.tables);
  }

  updateTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
  }: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
  }) {
    const table = this.tables.get(tableName);
    if (!table) {
      throw new Error("Table not found");
    }

    const newStart = start ? parseCellReference(start) : table.start;

    let newNumRows: SpreadsheetRangeEnd;
    if (numRows) {
      newNumRows = numRows;
    } else {
      if (table.endRow.type === "infinity") {
        newNumRows = table.endRow;
      } else {
        newNumRows = {
          type: "number",
          value: table.endRow.value - newStart.rowIndex,
        };
      }
    }

    const newTable = this.makeTable({
      tableName,
      sheetName: sheetName ?? table.sheetName,
      start: getCellReference(newStart),
      numRows: newNumRows,
      numCols: numCols ?? table.headers.size,
    });

    this.tables.set(tableName, newTable);

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();
    this.emit("tables-updated", this.tables);
  }

  removeTable({ tableName }: { tableName: string }) {
    const found = this.tables.delete(tableName);

    if (found) {
      // Re-evaluate all sheets since structured references might depend on this table
      this.reevaluate();
      this.triggerCellsUpdateEvent();
      this.emit("tables-updated", this.tables);
    }

    return found;
  }

  addSheet(name: string) {
    const sheet = {
      name,
      index: this.sheets.size,
      content: new Map(),
    };

    if (this.sheets.has(sheet.name)) {
      throw new Error("Sheet already exists");
    }

    this.sheets.set(name, sheet);

    // Emit sheet-added event
    this.emit("sheet-added", {
      sheetName: name,
    });
    return sheet;
  }

  removeSheet(sheetName: string) {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Remove the sheet
    this.sheets.delete(sheetName);

    // Clean up related data
    this.scopedNamedExpressions.delete(sheetName);
    this.tables.delete(sheetName);
    this.cellsUpdateListeners.delete(sheetName);

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    // Emit sheet-removed event
    this.emit("sheet-removed", {
      sheetName: sheetName,
    });

    return sheet;
  }

  renameSheet(sheetName: string, newName: string) {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    if (this.sheets.has(newName)) {
      throw new Error("Sheet with new name already exists");
    }

    // Update sheet name
    sheet.name = newName;

    // Update sheets map
    this.sheets.set(newName, sheet);
    this.sheets.delete(sheetName);

    // Update scoped named expressions
    const namedExpressions = this.scopedNamedExpressions.get(sheetName);
    if (namedExpressions) {
      this.scopedNamedExpressions.set(newName, namedExpressions);
      this.scopedNamedExpressions.delete(sheetName);
    }

    // Update tables that belong to the renamed sheet
    this.tables.forEach((table, tableName) => {
      if (table.sheetName === sheetName) {
        table.sheetName = newName;
      }
    });

    // Update cell update listeners
    const listeners = this.cellsUpdateListeners.get(sheetName);
    if (listeners) {
      this.cellsUpdateListeners.set(newName, listeners);
      this.cellsUpdateListeners.delete(sheetName);
    }

    // Update all formulas that reference this sheet
    this.sheets.forEach((sheet) => {
      sheet.content.forEach((cell, key) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          const formula = cell.slice(1);
          const updatedFormula = renameSheetInFormula(
            formula,
            sheetName,
            newName
          );

          // Only update if the formula actually changed
          if (updatedFormula !== formula) {
            sheet.content.set(key, `=${updatedFormula}`);
          }
        }
      });
    });

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    // Emit sheet-renamed event
    this.emit("sheet-renamed", {
      oldName: sheetName,
      newName: newName,
    });

    return sheet;
  }

  getTablesSerialized(): Map<string, TableDefinition> {
    return this.tables;
  }

  getNamedExpressionsSerialized(
    sheetName: string
  ): Map<string, NamedExpression> {
    return this.scopedNamedExpressions.get(sheetName) ?? new Map();
  }

  getGlobalNamedExpressionsSerialized(): Map<string, NamedExpression> {
    return this.globalNamedExpressions;
  }

  setNamedExpressions(
    sheetName: string,
    namedExpressions: Map<string, NamedExpression>
  ) {
    this.scopedNamedExpressions.set(sheetName, namedExpressions);
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  setGlobalNamedExpressions(namedExpressions: Map<string, NamedExpression>) {
    this.globalNamedExpressions = namedExpressions;
    this.reevaluate();
    this.triggerCellsUpdateEvent();
    this.emit("global-named-expressions-updated", namedExpressions);
  }

  setTables(tables: Map<string, TableDefinition>) {
    this.tables = tables;
    this.reevaluate();
    this.triggerCellsUpdateEvent();
    this.emit("tables-updated", this.tables);
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
