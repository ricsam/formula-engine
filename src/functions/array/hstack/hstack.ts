import {
  FormulaError,
  type AwaitingEvaluationResult,
  type CellAddress,
  type CellInRangeResult,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../../../core/types";
import { captureEvaluationErrors } from "../../../core/utils";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { getSuppliedArgument, type ASTNode } from "../../../parser/ast";
import {
  createSpillArea,
  getSpillWindow,
  outsideResultError,
  throwInfiniteCells,
  toArrayArgument,
  type ArrayArgument,
} from "../array-utils";

/**
 * HSTACK(array1, [array2], ...)
 * Appends arrays side by side, from left to right.
 *
 * The result has as many rows as the tallest argument and as many columns as
 * all arguments together. A shorter argument is padded with #N/A below its last
 * row; a single row or value is not repeated down. A single value is a 1x1
 * array, and so is an argument that evaluates to an error. A one-cell result is
 * returned as a single value.
 *
 * As in Excel, an omitted argument, as in `HSTACK(A1:A2,,B1:B2)`, makes the
 * result #VALUE!; pass `""` for a blank cell instead.
 *
 * Only the shapes of the arguments are needed to build the result: each of its
 * cells is read from the argument that holds it when the cell is read.
 *
 * Unbounded arguments stay lazy: `HSTACK(A:A, B1:B3)` has unbounded rows, with
 * #N/A below B3. Only the last argument may have an unbounded number of columns,
 * since no column could follow it; otherwise the result is #VALUE!.
 */
export const HSTACK: FunctionDefinition = {
  name: "HSTACK",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "HSTACK function takes at least 1 argument",
        errAddress: context.dependencyNode,
      };
    }

    const argNodes: ASTNode[] = [];
    for (let index = 0; index < node.args.length; index++) {
      const argNode = getSuppliedArgument(node, index);
      if (!argNode) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `HSTACK: array${index + 1} is omitted; use "" for a blank`,
          errAddress: context.dependencyNode,
        };
      }
      argNodes.push(argNode);
    }

    // Evaluate every argument before returning any of them, so that one pass
    // registers all dependencies.
    const results = argNodes.map((arg) => this.evaluateNode(arg, context));

    const args: ArrayArgument[] = [];
    /** The first result column of each argument. */
    const offsets: number[] = [];
    let rows = 0;
    let cols = 0;
    for (const [index, result] of results.entries()) {
      if (result.type === "awaiting-evaluation") {
        return result;
      }
      if (!Number.isFinite(cols)) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message:
            `HSTACK: array${index + 1} cannot follow array${index}, ` +
            "which has an unbounded number of columns",
          errAddress: context.dependencyNode,
        };
      }
      const arg = toArrayArgument(this, result, context);
      args.push(arg);
      offsets.push(cols);
      rows = Math.max(rows, arg.rows);
      cols += arg.cols;
    }

    /** Index of the argument that holds result column `x`. */
    const argumentAt = (x: number): number => {
      let low = 0;
      let high = offsets.length - 1;
      while (low < high) {
        const middle = (low + high + 1) >>> 1;
        if (offsets[middle]! <= x) {
          low = middle;
        } else {
          high = middle - 1;
        }
      }
      return low;
    };

    const padding = (index: number): ErrorEvaluationResult => ({
      type: "error",
      err: FormulaError.NA,
      message: `HSTACK: array${index + 1} has fewer rows than the result`,
      errAddress: context.dependencyNode,
    });

    const cellAt = (
      x: number,
      y: number,
      cellContext: EvaluationContext
    ): SingleEvaluationResult => {
      const index = argumentAt(x);
      const arg = args[index]!;
      if (y >= arg.rows) {
        return padding(index);
      }
      return arg.cellAt(x - offsets[index]!, y, cellContext);
    };

    if (rows === 1 && cols === 1) {
      return cellAt(0, 0, context);
    }

    const spillArea = (origin: CellAddress) =>
      createSpillArea(origin, rows, cols);

    return {
      type: "spilled-values",
      spillArea,
      source: "HSTACK function",
      evaluate: (spillOffset, cellContext) => {
        const { x, y } = spillOffset;
        if (x < 0 || y < 0 || x >= cols || y >= rows) {
          return outsideResultError("HSTACK", context);
        }
        return cellAt(x, y, cellContext);
      },
      evaluateAllCells: ({
        intersection,
        context: readContext,
        origin,
        lookupOrder,
      }) => {
        const window = getSpillWindow(spillArea(origin), origin, intersection);
        if (!window) {
          return { type: "values", values: [] };
        }

        const values: CellInRangeResult[] = [];
        let awaiting: AwaitingEvaluationResult | undefined;
        for (const [index, arg] of args.entries()) {
          const offset = offsets[index]!;
          // The argument's columns inside the window, as its own offsets.
          const x0 = Math.max(window.x0 - offset, 0);
          const x1 = Math.min(window.x1 - offset, arg.cols - 1);
          if (x0 > x1) {
            continue;
          }

          const y1 = Math.min(window.y1, arg.rows - 1);
          if (window.y0 <= y1) {
            if (arg.sparse) {
              const read = arg.readCells(readContext);
              if (read.type === "awaiting-evaluation") {
                awaiting ??= read;
              } else if (read.type !== "cells") {
                return read;
              } else {
                for (const cell of read.cells) {
                  if (
                    cell.x >= x0 &&
                    cell.x <= x1 &&
                    cell.y >= window.y0 &&
                    cell.y <= y1
                  ) {
                    values.push({
                      result: cell.result,
                      relativePos: { x: cell.x + offset, y: cell.y },
                    });
                  }
                }
              }
            } else {
              // A dense argument is bounded, so these loops end.
              for (let y = window.y0; y <= y1; y++) {
                for (let x = x0; x <= x1; x++) {
                  const result = captureEvaluationErrors(
                    readContext.dependencyNode,
                    () => arg.cellAt(x, y, readContext)
                  );
                  if (result.type === "awaiting-evaluation") {
                    awaiting ??= result;
                    continue;
                  }
                  values.push({ result, relativePos: { x: x + offset, y } });
                }
              }
            }
          }

          // #N/A padding below a shorter argument.
          const paddingStart = Math.max(window.y0, arg.rows);
          if (Number.isFinite(arg.rows) && paddingStart <= window.y1) {
            if (!Number.isFinite(window.y1) || !Number.isFinite(x1)) {
              throwInfiniteCells("HSTACK");
            }
            for (let y = paddingStart; y <= window.y1; y++) {
              for (let x = x0; x <= x1; x++) {
                values.push({
                  result: padding(index),
                  relativePos: { x: x + offset, y },
                });
              }
            }
          }
        }

        if (awaiting) {
          return awaiting;
        }
        values.sort(
          lookupOrder === "row-major"
            ? (a, b) =>
                a.relativePos.y - b.relativePos.y ||
                a.relativePos.x - b.relativePos.x
            : (a, b) =>
                a.relativePos.x - b.relativePos.x ||
                a.relativePos.y - b.relativePos.y
        );
        return { type: "values", values };
      },
    };
  },
};
