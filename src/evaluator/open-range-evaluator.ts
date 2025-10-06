import type { WorkbookManager } from "src/core/managers";
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
import { EvaluationError } from "./evaluation-error";

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
  *evaluateCellsInRange(options: {
    origin: RangeAddress;
    context: EvaluationContext;
  }): Iterable<EvaluateAllCellsResult> {
    const rangeNode = this.dependencyManager.getRangeNode(
      rangeAddressToKey(options.origin)
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
    const sheet = this.workbookManager.getSheet(options.origin);
    if (!sheet) {
      throw new EvaluationError(
        FormulaError.REF,
        `Sheet ${options.origin.sheetName} not found`
      );
    }

    const cellsInRange = this.workbookManager.getCellsInRange(options.origin);

    // Iterate over all defined cells in the sheet using optimized index-based iterator
    for (const address of cellsInRange) {
      const cellKey = cellAddressToKey(address);

      if (cellKey === options.context.originCell.key) {
        // if a cell in range for some reason is the origin, well, that's a cycle
        yield {
          result: {
            type: "error",
            err: FormulaError.CYCLE,
            message: "Cycle detected",
          },
          relativePos: {
            x:
              options.context.originCell.cellAddress.colIndex -
              options.origin.range.start.col,
            y:
              options.context.originCell.cellAddress.rowIndex -
              options.origin.range.start.row,
          },
        };
        return;
      }

      const cellNode = this.dependencyManager.getCellNode(cellKey);

      if (cellNode instanceof EmptyCellEvaluationNode) {
        throw new Error("A cell in range can not be an empty cell");
      }

      rangeNode.addDependency(cellNode);

      const result = cellNode.evaluationResult;

      if (result.type === "spilled-values") {
        const spilledResults = result.evaluateAllCells.call(this.evaluator, {
          context: rangeContext,
          evaluate: result.evaluate,
          intersection: options.origin.range,
          origin: address,
        });
        for (const spilledResult of spilledResults) {
          yield spilledResult;
        }
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

    const frontierCandidates = this.workbookManager.iterateFrontierCandidates(
      options.origin
    );

    // Evaluate frontier candidates first using the iterator
    for (const candidate of frontierCandidates) {
      const candidateKey = cellAddressToKey(candidate);
      const candidateNode = this.dependencyManager.getCellNode(candidateKey);
      const result = candidateNode.evaluationResult;

      if (candidateNode instanceof EmptyCellEvaluationNode) {
        throw new Error("A frontier dependencies can not be an empty cell");
      }

      rangeNode.addFrontierDependency(candidateNode);

      if (result) {
        if (result.type === "spilled-values") {
          const spillArea = result.spillArea(candidate);
          const intersects = checkRangeIntersection(
            spillArea,
            options.origin.range
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

            const spilledResults = Array.from(
              result.evaluateAllCells.call(this.evaluator, {
                context: rangeContext,
                evaluate: result.evaluate,
                intersection: options.origin.range,
                origin: candidate,
              })
            );
            for (const spilledResult of spilledResults) {
              yield spilledResult;
            }
          } else {
            rangeNode.maybeDiscardFrontierDependency(candidateNode); // downgraded!
          }
        } else {
          rangeNode.maybeDiscardFrontierDependency(candidateNode); // downgraded!
        }
      }
    }
  }
}
