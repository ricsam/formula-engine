import type { WorkbookManager } from "src/core/managers";
import {
  type CellAddress,
  type ErrorEvaluationResult,
  type EvaluationContext,
  type FunctionEvaluationResult,
  type LocalCellAddress,
  type SpilledValuesEvaluationResult,
  type SpilledValuesEvaluator,
  type SpreadsheetRange,
  type ValueEvaluationResult,
  FormulaError,
} from "src/core/types";
import {
  getCellReference,
  isCellInRange,
  parseCellReference,
} from "src/core/utils";
import { dependencyNodeToKey } from "src/core/utils/dependency-node-key";
import type { StoreManager } from "src/core/managers/store-manager";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

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
  }): Iterable<ValueEvaluationResult | ErrorEvaluationResult> {
    const rawContent = this.workbookManager.getSheet(options.origin)?.content;
    const { evaluate, context } = options;

    if (
      options.origin.sheetName === context.currentCell.sheetName &&
      isCellInRange(context.currentCell, options.origin.range)
    ) {
      yield {
        type: "error",
        err: FormulaError.CYCLE,
        message: "Cycle detected",
      };
      return;
    }

    if (!rawContent) {
      yield {
        type: "error",
        err: FormulaError.REF,
        message: `Sheet ${options.origin.sheetName} not found`,
      };
      return;
    }

    // Get frontier candidates that might spill into the range
    const frontierCandidates = this.workbookManager.getFrontierCandidates(
      options.origin.range,
      options.origin
    );

    // Keep track of cells we've already processed to avoid double-counting
    const processedCells = new Set<string>();

    // Evaluate frontier candidates first
    for (const candidate of frontierCandidates) {
      if (
        candidate.sheetName === context.currentCell.sheetName &&
        candidate.rowIndex === context.currentCell.rowIndex &&
        candidate.colIndex === context.currentCell.colIndex
      ) {
        continue;
      }

      const candidateKey = this.cellAddressToKey({
        rowIndex: candidate.rowIndex,
        colIndex: candidate.colIndex,
      });

      processedCells.add(candidateKey);

      const key = dependencyNodeToKey({
        address: candidate,
        sheetName: candidate.sheetName,
        workbookName: candidate.workbookName,
      });

      if (context.discardedFrontierDependencies.has(key)) {
        continue;
      }

      const result =
        this.storeManager.getEvaluatedNode(key)?.evaluationResult;

      context.frontierDependencies.add(key);

      if (result) {
        if (result.type === "spilled-values") {
          const spillArea = result.spillArea(candidate);
          const intersects = checkRangeIntersection(
            spillArea,
            options.origin.range
          );
          if (intersects) {
            context.dependencies.add(key);
            yield* this.handleSpilledValues({
              spillResult: result,
              targetRange: options.origin.range,
              candidate,
              context,
            });
          } else {
            context.discardedFrontierDependencies.add(key);
          }
        } else {
          context.discardedFrontierDependencies.add(key);
        }
      }
    }

    // Iterate over all defined cells in the sheet
    for (const address of this.iterateCellsInOpenRange(
      options.origin.range,
      rawContent
    )) {
      const cellKey = this.cellAddressToKey(address);

      // Skip if we already processed this cell as a frontier candidate
      if (processedCells.has(cellKey)) {
        continue;
      }

      // const offsetLeft = address.colIndex - range.start.col;
      // const offsetTop = address.rowIndex - range.start.row;

      const result = this.storeManager.evalTimeSafeEvaluateCell(
        {
          ...address,
          sheetName: options.origin.sheetName,
          workbookName: options.origin.workbookName,
        },
        context
      );

      if (result?.type === "spilled-values") {
        const spillHandleResult = this.handleSpilledValues({
          spillResult: result,
          targetRange: options.origin.range,
          context,
          candidate: {
            ...address,
            sheetName: options.origin.sheetName,
            workbookName: options.origin.workbookName,
          },
        });
        yield* spillHandleResult;
      } else {
        yield result ?? {
          type: "error",
          err: FormulaError.REF,
          message: `Error evaluating cell ${cellKey} #2`,
        };
      }
    }
  }

  /**
   * Iterator for cells within an open range
   */
  private *iterateCellsInOpenRange(
    range: SpreadsheetRange,
    rawContent: Map<string, any>
  ): Generator<LocalCellAddress> {
    for (const [key, value] of rawContent) {
      const { rowIndex, colIndex } = parseCellReference(key);

      // Check if cell is within range bounds
      if (rowIndex < range.start.row || colIndex < range.start.col) continue;

      if (range.end.row.type === "number" && rowIndex > range.end.row.value)
        continue;
      if (range.end.col.type === "number" && colIndex > range.end.col.value)
        continue;

      yield {
        rowIndex,
        colIndex,
      };
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
  }): Iterable<ValueEvaluationResult | ErrorEvaluationResult> {
    const spillArea = options.spillResult.spillArea(options.candidate);

    // Calculate intersection first
    const intersection = getRangeIntersection(spillArea, options.targetRange);

    if (!intersection) {
      yield {
        type: "error",
        err: FormulaError.REF,
        message: "Intersection is not valid #3",
      };
      return;
    }

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

  /**
   * Convert a cell address to a string key
   */
  private cellAddressToKey(addr: LocalCellAddress): string {
    return getCellReference(addr);
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
