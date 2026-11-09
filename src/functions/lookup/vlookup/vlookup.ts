import {
  FormulaError,
  type CellValue,
  type EvaluateAllCellsResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpreadsheetRange,
} from "../../../core/types";
import {
  createBroadcastResult,
  toBroadcastOperand,
} from "../../../evaluator/broadcast";
import { getSpillAreaShape } from "../../../evaluator/array-shape";
import { getSuppliedArgument } from "../../../parser/ast";
import {
  blankResult,
  readLogicalOption,
  toArrayArgument,
} from "../../array/array-utils";
import { matchOperation } from "../match/match";

const ZERO: FunctionEvaluationResult = {
  type: "value",
  result: { type: "number", value: 0 },
};

function describeValue(value: CellValue): string {
  return value.type === "string" && value.value === ""
    ? "an empty value"
    : value.type;
}

/**
 * VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
 *
 * Looks for `lookup_value` in the first column of `table_array` and returns the
 * cell in the same row of column `col_index_num` (1 is the first column).
 *
 * Examples:
 *   VLOOKUP("P-102", A2:D50, 3, FALSE) - column C of the row whose A is P-102
 *   VLOOKUP(E2, A:B, 2, FALSE) - an open-ended table
 *   VLOOKUP(72, {0,"F";60,"D";70,"C";80,"B";90,"A"}, 2) - "C", the band 72 is in
 *
 * - The first column is searched exactly like MATCH searches its lookup_array:
 *   `range_lookup` FALSE (or 0) is MATCH's exact match type 0 and TRUE (or any
 *   other number) its type 1, the largest value less than or equal to
 *   `lookup_value` in an ascending column. So, as in MATCH, text is compared
 *   without regard to case, values only match values of the same type, and
 *   `lookup_value` must be a number or text. Wildcards (`*`, `?`) are not
 *   supported, as in MATCH; Excel supports them in an exact VLOOKUP.
 * - `range_lookup` defaults to TRUE when it is left out, as in Excel, and an
 *   omitted one (a trailing comma) is FALSE: `VLOOKUP(x, t, 2,)` is exact. An
 *   empty cell is FALSE too; text gives #VALUE!.
 * - `col_index_num` must be a number and is truncated; below 1 it gives
 *   #VALUE! and beyond the last column of `table_array` #REF!, as in Excel,
 *   whether or not the value is found. An omitted `col_index_num` is 0.
 * - No match gives #N/A.
 * - A blank cell in the returned column is returned as `""`, where Excel shows
 *   0. Errors in the table are returned when their cell is chosen, and errors
 *   in the arguments propagate.
 * - An array `lookup_value` or `col_index_num` looks up every element and
 *   spills; the two are broadcast against each other like the operands of an
 *   operator. `VLOOKUP(E2:E5, A:C, 3, FALSE)` spills four results and
 *   `VLOOKUP(E2, A:C, {2,3}, FALSE)` two columns of one row.
 * - A single lookup reads the first column through a range of its own and then
 *   only the chosen cell, so it depends on that column and that cell rather
 *   than on the whole table. Several lookups read the table once.
 */
export const VLOOKUP: FunctionDefinition = {
  name: "VLOOKUP",
  evaluate: function (node, context): FunctionEvaluationResult {
    const error = (
      err: FormulaError,
      message: string
    ): SingleEvaluationResult => ({
      type: "error",
      err,
      message: `VLOOKUP: ${message}`,
      errAddress: context.dependencyNode,
    });

    if (node.args.length < 3 || node.args.length > 4) {
      return error(FormulaError.VALUE, "takes 3 or 4 arguments");
    }
    const tableNode = getSuppliedArgument(node, 1);
    if (!tableNode) {
      return error(FormulaError.VALUE, "table_array is required");
    }

    // Evaluate every argument before returning any of them, so that one pass
    // registers all dependencies. Omitted values are 0, as in Excel.
    const lookupNode = getSuppliedArgument(node, 0);
    const columnNode = getSuppliedArgument(node, 2);
    const rangeLookupNode = getSuppliedArgument(node, 3);
    const lookupResult = lookupNode
      ? this.evaluateNode(lookupNode, context)
      : ZERO;
    const tableResult = this.evaluateNode(tableNode, context);
    const columnResult = columnNode
      ? this.evaluateNode(columnNode, context)
      : ZERO;
    const rangeLookupResult =
      rangeLookupNode && this.evaluateNode(rangeLookupNode, context);

    for (const result of [
      lookupResult,
      tableResult,
      columnResult,
      rangeLookupResult,
    ]) {
      if (result?.type === "awaiting-evaluation") {
        return result;
      }
    }
    if (lookupResult.type === "error") {
      return lookupResult;
    }
    if (
      tableResult.type === "error" ||
      tableResult.type === "awaiting-evaluation"
    ) {
      return tableResult;
    }
    if (columnResult.type === "error") {
      return columnResult;
    }

    // Left out, range_lookup is TRUE; omitted (a trailing comma), FALSE.
    let approximate = node.args.length < 4;
    if (rangeLookupResult) {
      if (rangeLookupResult.type === "awaiting-evaluation") {
        return rangeLookupResult;
      }
      const value = readLogicalOption(
        rangeLookupResult,
        "VLOOKUP range_lookup",
        context
      );
      if (typeof value !== "boolean") {
        return value;
      }
      approximate = value;
    }

    // The table's shape and its first column, positioned from the table's
    // top-left cell.
    let cols = 1;
    let firstColumn: EvaluateAllCellsResult;
    if (tableResult.type === "spilled-values") {
      const area = tableResult.spillArea(context.cellAddress);
      cols = getSpillAreaShape(area).cols;
      const firstColumnArea: SpreadsheetRange = {
        start: area.start,
        end: {
          col: { type: "number", value: area.start.col },
          row: area.end.row,
        },
      };
      firstColumn = tableResult.evaluateAllCells.call(this, {
        context,
        origin: context.cellAddress,
        evaluate: tableResult.evaluate,
        lookupOrder: "col-major",
        intersection: firstColumnArea,
      });
      if (firstColumn.type !== "values") {
        return firstColumn;
      }
    } else {
      firstColumn = {
        type: "values",
        values: [{ result: tableResult, relativePos: { x: 0, y: 0 } }],
      };
    }

    /** Looks up one value; `cellAt` reads a cell of the table. */
    const lookup = (
      lookupValue: CellValue,
      columnValue: CellValue,
      cellAt: (x: number, y: number) => SingleEvaluationResult
    ): SingleEvaluationResult => {
      if (columnValue.type === "infinity") {
        return columnValue.sign === "positive"
          ? error(FormulaError.REF, "col_index_num is beyond table_array")
          : error(FormulaError.VALUE, "col_index_num must be at least 1");
      }
      if (columnValue.type !== "number") {
        return error(
          FormulaError.VALUE,
          `col_index_num must be a number, got ${describeValue(columnValue)}`
        );
      }
      const column = Math.trunc(columnValue.value);
      if (column < 1) {
        return error(
          FormulaError.VALUE,
          `col_index_num must be at least 1, got ${column}`
        );
      }
      if (column > cols) {
        return error(
          FormulaError.REF,
          `col_index_num ${column} is beyond the ${cols} columns of table_array`
        );
      }

      const position = matchOperation(
        lookupValue,
        firstColumn,
        approximate ? 1 : 0,
        context,
        false
      );
      if (position.type === "error") {
        return { ...position, message: `VLOOKUP: ${position.message}` };
      }
      if (position.type !== "value" || position.result.type !== "number") {
        return error(FormulaError.NA, "lookup_value was not found");
      }
      return cellAt(column - 1, position.result.value - 1);
    };

    if (lookupResult.type === "value" && columnResult.type === "value") {
      // Only the chosen cell is read, with its own dependency.
      return lookup(lookupResult.result, columnResult.result, (x, y) =>
        tableResult.type === "spilled-values"
          ? tableResult.evaluate({ x, y }, context)
          : tableResult
      );
    }

    // Several lookups: read the table once and look every element up in it.
    const table = toArrayArgument(this, tableResult, context);
    const read = table.readCells(context);
    if (read.type !== "cells") {
      return read;
    }
    const key = (x: number, y: number) =>
      Number.isFinite(cols) ? y * cols + x : `${x},${y}`;
    const cells = new Map<number | string, SingleEvaluationResult>();
    for (const cell of read.cells) {
      cells.set(key(cell.x, cell.y), cell.result);
    }

    return createBroadcastResult({
      operands: [
        toBroadcastOperand(lookupResult, context.cellAddress),
        toBroadcastOperand(columnResult, context.cellAddress),
      ],
      context,
      source: "VLOOKUP with an array argument",
      combine: (operandCell) => {
        const lookupValue = operandCell(0);
        if (lookupValue.type !== "value") {
          return lookupValue;
        }
        const columnValue = operandCell(1);
        if (columnValue.type !== "value") {
          return columnValue;
        }
        return lookup(
          lookupValue.result,
          columnValue.result,
          (x, y) => cells.get(key(x, y)) ?? blankResult()
        );
      },
    });
  },
};
