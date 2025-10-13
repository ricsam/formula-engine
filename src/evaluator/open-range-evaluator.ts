import type { LookupOrder, WorkbookManager } from "src/core/managers";
import {
  type CellAddress,
  type EvaluateAllCellsResult,
  type RangeAddress,
  type SpilledValuesEvaluationResult,
  type SpilledValuesEvaluator,
  type SpreadsheetRange,
  FormulaError,
} from "src/core/types";
import {
  cellAddressToKey,
  getCellReference,
  isCellInRange,
  checkRangeIntersection,
  rangeAddressToKey,
} from "src/core/utils";

import type { DependencyManager } from "src/core/managers/dependency-manager";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import { EvaluationContext } from "./evaluation-context";
import { flags } from "src/debug/flags";
import type { CellEvalNode } from "./cell-eval-node";
import { EmptyCellEvaluationNode } from "./empty-cell-evaluation-node";
import { AwaitingEvaluationError, EvaluationError } from "./evaluation-error";

/**
 * Utility class for evaluating cells within open-ended ranges
 */
export class OpenRangeEvaluator {
  constructor(
    private dependencyManager: DependencyManager,
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
  evaluateCellsInRange(options: {
    address: RangeAddress;
    context: EvaluationContext;
    lookupOrder: LookupOrder;
  }): EvaluateAllCellsResult[] {
    const results: EvaluateAllCellsResult[] = [];
    const rangeNode = this.dependencyManager.getRangeNode(
      rangeAddressToKey(options.address)
    );

    if (options.context.originCell instanceof EmptyCellEvaluationNode) {
      throw new Error(
        "If the origin cell is an empty cell, we could not be evaluating a range"
      );
    }
    options.context.originCell.addDependency(rangeNode);
    const rangeContext = new EvaluationContext(
      rangeNode,
      options.context.originCell
    );

    // Check if the sheet exists
    const sheet = this.workbookManager.getSheet(options.address);
    if (!sheet) {
      throw new EvaluationError(
        FormulaError.REF,
        `Sheet ${options.address.sheetName} not found`
      );
    }

    const evalOrder = rangeNode.getRangeEvalOrder();

    for (const entry of evalOrder) {
      if (entry.type === "value") {
        const entryAddress = entry.address;
        const cellKey = cellAddressToKey(entryAddress);

        const cellNode = this.dependencyManager.getCellNode(cellKey);
        const result = cellNode.evaluationResult;
        if (result.type === "awaiting-evaluation") {
          throw new AwaitingEvaluationError(
            options.context.originCell.cellAddress,
            entryAddress
          );
        }

        const relativePos = {
          x: entryAddress.colIndex - options.address.range.start.col,
          y: entryAddress.rowIndex - options.address.range.start.row,
        };

        if (result.type === "spilled-values") {
          const spilledResult = result.evaluate({ x: 0, y: 0 }, rangeContext);
          results.push({
            result: spilledResult,
            relativePos,
          });
        } else {
          results.push({ result: result, relativePos });
        }
      } else if (entry.type === "empty_cell" || entry.type === "empty_range") {
        for (const candidate of entry.candidates) {
          const candidateKey = cellAddressToKey(candidate);
          const candidateNode =
            this.dependencyManager.getCellNode(candidateKey);
          const result = candidateNode.evaluationResult;

          if (candidateNode instanceof EmptyCellEvaluationNode) {
            throw new Error("A frontier dependency can not be an empty cell");
          }

          if (result.type === "spilled-values") {
            const spillArea = result.spillArea(candidate);
            if (entry.type === "empty_range") {
              const intersects = checkRangeIntersection(
                spillArea,
                entry.address.range
              );
              if (intersects) {
                rangeNode.maybeUpgradeFrontierDependency(candidateNode); // upgraded!
                // When a spilled range intersects with our target range, we need to evaluate
                // only the cells that fall within the intersection area.
                //
                // Example: If cell A10 contains a spilled range that covers A10:B11,
                // and our target range is B10:INFINITY, then we only want to evaluate
                // the intersection B10:B11 from the spilled range.
                //
                // The evaluateAllCells method expects the intersection to be passed
                // so it can limit evaluation to only the relevant cells.

                const spilledResults = result.evaluateAllCells.call(
                  this.evaluator,
                  {
                    context: rangeContext,
                    evaluate: result.evaluate,
                    intersection: entry.address.range,
                    origin: candidate,
                    lookupOrder: options.lookupOrder,
                  }
                );
                results.push(...spilledResults);
              } else {
                rangeNode.maybeDiscardFrontierDependency(candidateNode); // downgraded!
              }
            } else {
              const intersects = isCellInRange(entry.address, spillArea);
              if (intersects) {
                rangeNode.maybeUpgradeFrontierDependency(candidateNode); // upgraded!
                // When a spilled range intersects with our target range, we need to evaluate
                // only the cells that fall within the intersection area.
                //
                // Example: If cell A10 contains a spilled range that covers A10:B11,
                // and our target range is B10:INFINITY, then we only want to evaluate
                // the intersection B10:B11 from the spilled range.
                //
                // The evaluateAllCells method expects the intersection to be passed
                // so it can limit evaluation to only the relevant cells.

                const relativePos = {
                  x: entry.address.colIndex - candidate.colIndex,
                  y: entry.address.rowIndex - candidate.rowIndex,
                };
                const spilledResult = result.evaluate(
                  relativePos,
                  rangeContext
                );

                results.push({
                  relativePos: {
                    x: entry.address.colIndex - options.address.range.start.col,
                    y: entry.address.rowIndex - options.address.range.start.row,
                  },
                  result: spilledResult,
                });
              } else {
                rangeNode.maybeDiscardFrontierDependency(candidateNode); // downgraded!
              }
            }
          } else {
            rangeNode.maybeDiscardFrontierDependency(candidateNode); // downgraded!
          }
        }
      }
    }

    return results;
  }
}
