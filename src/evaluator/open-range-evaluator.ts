import type { WorkbookManager } from "src/core/managers";
import {
  type CellAddress,
  type EvaluateAllCellsResult,
  type SpilledValuesEvaluationResult,
  type SpilledValuesEvaluator,
  type SpreadsheetRange,
  FormulaError,
} from "src/core/types";
import {
  cellAddressToKey,
  getCellReference,
  isCellInRange,
} from "src/core/utils";

import type { StoreManager } from "src/core/managers/store-manager";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { EvaluationContext } from "./evaluation-context";

/**
 * Utility class for evaluating cells within open-ended ranges
 */
export class OpenRangeEvaluator {
  constructor(
    private storeManager: StoreManager,
    private workbookManager: WorkbookManager,
    private evaluator: FormulaEvaluator
  ) {}

  /**
   * Evaluates all cells within an open-ended range and returns their values
   * @param range - The spreadsheet range (may have infinite bounds)
   * @param sheetName - The sheet to evaluate on
   * @param context - Evaluation context
   * @returns Array of evaluation results or INFINITY if infinite spill detected
   */
  *evaluateCellsInRange(options: {
    origin: {
      range: SpreadsheetRange;
      sheetName: string;
      workbookName: string;
    };
    context: EvaluationContext;
    evaluate: SpilledValuesEvaluator;
  }): Iterable<EvaluateAllCellsResult> {
    const { evaluate, context } = options;

    if (
      options.origin.sheetName === context.currentCell.sheetName &&
      isCellInRange(context.currentCell, options.origin.range)
    ) {
      yield {
        result: {
          type: "error",
          err: FormulaError.CYCLE,
          message: "Cycle detected",
        },
        relativePos: {
          x: context.currentCell.colIndex - options.origin.range.start.col,
          y: context.currentCell.rowIndex - options.origin.range.start.row,
        },
      };
      return;
    }

    // Check if the sheet exists
    const sheet = this.workbookManager.getSheet(options.origin);
    if (!sheet) {
      yield {
        result: {
          type: "error",
          err: FormulaError.REF,
          message: `Sheet ${options.origin.sheetName} not found`,
        },
        relativePos: {
          x: context.currentCell.colIndex - options.origin.range.start.col,
          y: context.currentCell.rowIndex - options.origin.range.start.row,
        },
      };
      return;
    }

    // todo change to using the iterator
    // const frontierCandidates = this.workbookManager.getFrontierCandidates(
    //   options.origin.range,
    //   options.origin
    // );
    const frontierCandidates = this.workbookManager.iterateFrontierCandidates(
      options.origin.range,
      options.origin
    );

    // Evaluate frontier candidates first using the iterator
    for (const candidate of frontierCandidates) {
      const key = cellAddressToKey(candidate);

      if (context.isFrontierDependencyDiscarded(key, options.origin.range)) {
        continue;
      }

      const result = this.storeManager.getEvaluatedNode(key)?.evaluationResult;

      context.addFrontierDependency(key, options.origin.range);

      if (result) {
        if (result.type === "spilled-values") {
          const spillArea = result.spillArea(candidate);
          const intersects = checkRangeIntersection(
            spillArea,
            options.origin.range
          );
          if (intersects) {
            context.maybeUpgradeFrontierDependency(key, options.origin.range); // upgraded!
            yield* this.handleSpilledValues({
              spillResult: result,
              targetRange: options.origin.range,
              candidate,
              context,
            });
          } else {
            context.maybeDiscardFrontierDependency(key, options.origin.range); // downgraded!
          }
        } else {
          context.maybeDiscardFrontierDependency(key, options.origin.range); // downgraded!
        }
      }
    }

    // todo change to using the iterator
    // const cellsInRange = this.workbookManager.getCellsInRange(
    //   options.origin,
    //   options.origin.range
    // );
    const cellsInRange = this.workbookManager.iterateCellsInRange(
      options.origin,
      options.origin.range
    );

    // Iterate over all defined cells in the sheet using optimized index-based iterator
    for (const address of cellsInRange) {
      const cellKey = getCellReference(address);

      const result = this.storeManager.evalTimeSafeEvaluateCell(
        {
          ...address,
          sheetName: options.origin.sheetName,
          workbookName: options.origin.workbookName,
        },
        context
      );

      if (result?.type === "spilled-values") {
        const candidate: CellAddress = {
          ...address,
          sheetName: options.origin.sheetName,
          workbookName: options.origin.workbookName,
        };
        const spillHandleResult = this.handleSpilledValues({
          spillResult: result,
          targetRange: options.origin.range,
          context,
          candidate,
        });
        yield* spillHandleResult;
      } else {
        const relativePos = {
          x: address.colIndex - options.origin.range.start.col,
          y: address.rowIndex - options.origin.range.start.row,
        };
        yield result
          ? { result: result, relativePos }
          : {
              result: {
                type: "error",
                err: FormulaError.REF,
                message: `Error evaluating cell ${cellKey} #2`,
              },
              relativePos,
            };
      }
    }
  }

  /**
   * Handle spilled values that may intersect with the target range
   */
  *handleSpilledValues(options: {
    spillResult: SpilledValuesEvaluationResult;
    targetRange: SpreadsheetRange;
    candidate: CellAddress;
    context: EvaluationContext;
  }): Iterable<EvaluateAllCellsResult> {
    // When a spilled range intersects with our target range, we need to evaluate
    // only the cells that fall within the intersection area.
    //
    // Example: If cell A10 contains a spilled range that covers A10:B11,
    // and our target range is B10:INFINITY, then we only want to evaluate
    // the intersection B10:B11 from the spilled range.
    //
    // The evaluateAllCells method expects the intersection to be passed
    // so it can limit evaluation to only the relevant cells.

    return yield* options.spillResult.evaluateAllCells.call(this.evaluator, {
      context: options.context,
      evaluate: options.spillResult.evaluate,
      intersection: options.targetRange,
      origin: options.candidate,
    });
  }
}

/**
 * Check if two ranges intersect
 */
export function checkRangeIntersection(
  range1: SpreadsheetRange,
  range2: SpreadsheetRange
): boolean {
  // Check if ranges don't intersect
  if (
    range1.end.col.type === "number" &&
    range2.start.col > range1.end.col.value
  )
    return false;
  if (
    range2.end.col.type === "number" &&
    range1.start.col > range2.end.col.value
  )
    return false;
  if (
    range1.end.row.type === "number" &&
    range2.start.row > range1.end.row.value
  )
    return false;
  if (
    range2.end.row.type === "number" &&
    range1.start.row > range2.end.row.value
  )
    return false;

  return true;
}

/**
 * Get the intersection of two ranges
 */
export function getRangeIntersection(
  range1: SpreadsheetRange,
  range2: SpreadsheetRange
): SpreadsheetRange | null {
  if (!checkRangeIntersection(range1, range2)) {
    return null;
  }

  const startRow = Math.max(range1.start.row, range2.start.row);
  const startCol = Math.max(range1.start.col, range2.start.col);

  let endRow, endCol;

  // Handle end row
  if (
    range1.end.row.type === "infinity" &&
    range2.end.row.type === "infinity"
  ) {
    endRow = { type: "infinity" as const, sign: "positive" as const };
  } else if (
    range1.end.row.type === "number" &&
    range2.end.row.type === "number"
  ) {
    endRow = {
      type: "number" as const,
      value: Math.min(range1.end.row.value, range2.end.row.value),
    };
  } else {
    // One is finite, one is infinite
    endRow = range1.end.row.type === "number" ? range1.end.row : range2.end.row;
  }

  // Handle end col
  if (
    range1.end.col.type === "infinity" &&
    range2.end.col.type === "infinity"
  ) {
    endCol = { type: "infinity" as const, sign: "positive" as const };
  } else if (
    range1.end.col.type === "number" &&
    range2.end.col.type === "number"
  ) {
    endCol = {
      type: "number" as const,
      value: Math.min(range1.end.col.value, range2.end.col.value),
    };
  } else {
    // One is finite, one is infinite
    endCol = range1.end.col.type === "number" ? range1.end.col : range2.end.col;
  }

  return {
    start: { row: startRow, col: startCol },
    end: { row: endRow, col: endCol },
  };
}
