import { flags } from "src/debug/flags";
import { AstEvaluationNode } from "src/evaluator/dependency-nodes/ast-evaluation-node";
import { CellValueNode } from "src/evaluator/dependency-nodes/cell-value-node";
import { EvaluationContext } from "src/evaluator/evaluation-context";
import {
  EvaluationError,
  SheetNotFoundError,
} from "src/evaluator/evaluation-error";
import { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellInRangeResult,
  type CellValue,
  type ErrorEvaluationResult,
  type EvaluateAllCellsResult,
  type EvaluationOrder,
  type FunctionEvaluationResult,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type TableDefinition,
  type ValueEvaluationResult,
} from "../types";
import {
  captureEvaluationErrors,
  cellAddressToKey,
  checkRangeIntersection,
  getCellReference,
  isCellInRange,
  isRangeOneCell,
  keyToCellAddress,
  parseCellReference,
} from "../utils";
import type { DependencyManager } from "./dependency-manager";
import type { WorkbookManager } from "./workbook-manager";
import { SpillMetaNode } from "src/evaluator/dependency-nodes/spill-meta-node";
import { EmptyCellEvaluationNode } from "src/evaluator/dependency-nodes/empty-cell-evaluation-node";
import type { TableManager } from "./table-manager";

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private tableManager: TableManager,
    private formulaEvaluator: FormulaEvaluator,
    private dependencyManager: DependencyManager
  ) {}

  clearEvaluationCache(): void {
    this.dependencyManager.clearEvaluationCache();
  }

  evaluationResultToSerializedValue(
    evaluation: SingleEvaluationResult,
    cellAddress: CellAddress,
    debug?: boolean
  ): SerializedCellValue {
    if (
      evaluation.type !== "error" &&
      evaluation.type !== "awaiting-evaluation"
    ) {
      const value = evaluation.result;

      return value.type === "infinity"
        ? value.sign === "positive"
          ? "INFINITY"
          : "-INFINITY"
        : value.value;
    }

    if (evaluation.type === "awaiting-evaluation") {
      return (
        evaluation.errAddress.key +
        " is awaiting evaluation of " +
        evaluation.waitingFor.key
      );
    }

    if (debug) {
      const errAddress = evaluation.errAddress.key;
      if (errAddress === cellAddressToKey(cellAddress)) {
        return evaluation.err + " " + evaluation.message;
      }
      return (
        evaluation.err +
        " in " +
        evaluation.errAddress.key +
        " " +
        evaluation.message
      );
    }

    return evaluation.err;
  }

  evaluateEmptyCell(node: EmptyCellEvaluationNode): void {
    node.resetDirectDepsUpdated();

    if (node.resolved) {
      const result = node.evaluationResult;
      if (result && result.type !== "awaiting-evaluation") {
        return;
      }
    }

    const ctx = new EvaluationContext(
      this.tableManager,
      node,
      node.cellAddress
    );
    const inSpilled = this.dependencyManager.getSpillValue(node.cellAddress);

    if (inSpilled) {
      // if we are spilling then we can just add the spill origin as a dependency and evaluate the spilled value
      const spillTarget = this.dependencyManager.getSpilledAddress(
        node.cellAddress,
        inSpilled
      );
      const spillOriginKey = cellAddressToKey(inSpilled.origin).replace(
        /^[^:]+:/,
        "spill-meta:"
      );
      const spillMetaNode =
        this.dependencyManager.getSpillMetaNode(spillOriginKey);
      node.addDependency(spillMetaNode);
      const result = spillMetaNode.evaluationResult;
      if (result.type === "spilled-values") {
        // let's evaluate the spilled value to extract dependencies
        const evaluation = captureEvaluationErrors(spillMetaNode, () => {
          return result.evaluate(spillTarget.spillOffset, ctx);
        });
        node.setEvaluationResult(evaluation);
      }
    } else {
      // upgrade any frontier dependencies that spill into the range
      node.upgradeFrontierDependencies();

      const evaluationResult: SingleEvaluationResult = {
        type: "value",
        result: this.convertScalarValueToCellValue(""),
      };
      // for now let's just store the empty value, the next time the cell is evaluated isSpilled will be true and the spilled value will be evaluated
      node.setEvaluationResult(evaluationResult);
    }
  }

  evaluateRangeNode(node: RangeEvaluationNode): void {
    if (node.resolved) {
      return;
    }

    node.resetDirectDepsUpdated();

    const result = captureEvaluationErrors(node, (): EvaluateAllCellsResult => {
      node.upgradeFrontierDependencies();

      const evalOrder = node.getRangeEvalOrder();

      const results: CellInRangeResult[] = [];

      for (const entry of evalOrder) {
        if (entry.type === "value") {
          const entryAddress = entry.address;
          const result = entry.node.evaluationResult;
          if (result.type === "awaiting-evaluation") {
            return {
              type: "awaiting-evaluation",
              waitingFor: entry.node,
              errAddress: node,
            };
          }

          const relativePos = {
            x: entryAddress.colIndex - node.address.range.start.col,
            y: entryAddress.rowIndex - node.address.range.start.row,
          };

          results.push({ result: result, relativePos });
        } else if (
          entry.type === "empty_cell" ||
          entry.type === "empty_range"
        ) {
          for (const candidateNode of entry.candidates) {
            if (candidateNode.evaluationResult.type === "spilled-values") {
              const spillArea = candidateNode.evaluationResult.spillArea(
                candidateNode.cellAddress
              );
              if (entry.type === "empty_range") {
                const intersects = checkRangeIntersection(
                  spillArea,
                  entry.address.range
                );
                if (intersects) {
                  // When a spilled range intersects with our target range, we need to evaluate
                  // only the cells that fall within the intersection area.
                  //
                  // Example: If cell A10 contains a spilled range that covers A10:B11,
                  // and our target range is B10:INFINITY, then we only want to evaluate
                  // the intersection B10:B11 from the spilled range.
                  //
                  // The evaluateAllCells method expects the intersection to be passed
                  // so it can limit evaluation to only the relevant cells.
                  const ctx = new EvaluationContext(
                    this.tableManager,
                    node,
                    candidateNode.cellAddress
                  );
                  const spilledResults =
                    candidateNode.evaluationResult.evaluateAllCells.call(
                      this.formulaEvaluator,
                      {
                        context: ctx,
                        evaluate: candidateNode.evaluationResult.evaluate,
                        intersection: entry.address.range,
                        origin: candidateNode.cellAddress,
                        lookupOrder: "col-major",
                      }
                    );
                  if (spilledResults.type === "values") {
                    results.push(...spilledResults.values);
                  } else {
                    return spilledResults;
                  }
                }
              } else {
                const intersects = isCellInRange(entry.address, spillArea);
                if (intersects) {
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
                    x:
                      entry.address.colIndex -
                      candidateNode.cellAddress.colIndex,
                    y:
                      entry.address.rowIndex -
                      candidateNode.cellAddress.rowIndex,
                  };
                  const ctx = new EvaluationContext(
                    this.tableManager,
                    node,
                    candidateNode.cellAddress
                  );
                  const spilledResult = candidateNode.evaluationResult.evaluate(
                    relativePos,
                    ctx
                  );

                  results.push({
                    relativePos: {
                      x: entry.address.colIndex - node.address.range.start.col,
                      y: entry.address.rowIndex - node.address.range.start.row,
                    },
                    result: spilledResult,
                  });
                }
              }
            }
          }
        }
      }

      return {
        type: "values",
        values: results,
      };
    });

    node.setResult(result);
  }

  evaluateCellNode(node: CellValueNode | SpillMetaNode): void {
    // Enable caching for resolved nodes
    if (node.resolved) {
      return;
    }

    node.resetDirectDepsUpdated();

    const ctx = new EvaluationContext(
      this.tableManager,
      node,
      node.cellAddress
    );

    if (node instanceof CellValueNode && node.spillMeta) {
      // we are evaluating a e.g. A1 in A1=SEQUENCE(10), where we want the value in the cell in A1, i.e. 1
      // As A1 is already pointing to a spill meta node, everything has been setup already,
      // we just need to evaluate the spill origin and assign the result to the currentDepNode
      const spillOrigin = node.spillMeta;
      if (spillOrigin.evaluationResult.type === "spilled-values") {
        const result = spillOrigin.evaluationResult.evaluate(
          { x: 0, y: 0 },
          ctx
        );
        node.setEvaluationResult(result);
        return;
      }
    }

    let content: SerializedCellValue;
    try {
      content = this.workbookManager.getSerializedCellValue(node.cellAddress);
    } catch (err) {
      const evaluationResult: ErrorEvaluationResult = {
        type: "error",
        err: FormulaError.ERROR,
        message: "Syntax error",
        errAddress: node,
      };
      node.setEvaluationResult(evaluationResult);
      return;
    }

    if (typeof content !== "string" || !content.startsWith("=")) {
      if (node instanceof SpillMetaNode) {
        node.setEvaluationResult({
          type: "does-not-spill",
        });
        return;
      }
      // Static value cells cannot have frontier dependencies
      const result: ValueEvaluationResult = {
        type: "value",
        result: this.convertScalarValueToCellValue(content),
      };
      node.setEvaluationResult(result);
      return;
    }

    let evaluation: FunctionEvaluationResult =
      this.formulaEvaluator.evaluateFormula(content.slice(1), ctx);

    if (node.cellAddress.colIndex === 5 && node.cellAddress.rowIndex === 0 && content.includes("SEQUENCE")) {
      console.log(`\n🔍 F1 evaluation:`);
      console.log(`  Formula: ${content}`);
      console.log(`  Evaluation type: ${evaluation.type}`);
    }

    // the evaluated cell IS A spilling formula, e.g. if dependencyKey points to A1, then the formula is e.g. A1=SEQUENCE(10), or A1=A3:B5
    if (evaluation.type === "spilled-values") {
      const spillArea = evaluation.spillArea(node.cellAddress);
      
      // Check if spill area is blocked (but allow single-cell "spills")
      if (!isRangeOneCell(spillArea)) {
        if (!this.canSpill(node.cellAddress, spillArea)) {
          // Override evaluation with SPILL error, but continue execution to set up nodes
          evaluation = {
            type: "error",
            err: FormulaError.SPILL,
            message: "Cannot spill - area is blocked",
            errAddress: node,
          };
        } else {
          // Spill succeeds - register it
          this.dependencyManager.setSpilledValue(node.key, {
            spillOnto: spillArea,
            origin: node.cellAddress,
          });
        }
      }
      
      // Set up spill meta node and evaluation results (even if spill failed)
      if (node instanceof SpillMetaNode) {
        // we have already setup an origin/spill meta node relationship,
        // so we are just reevaluating the spill meta node here
        node.setEvaluationResult(evaluation);
      } else {
        const spillMetaNode = this.dependencyManager.getSpillMetaNode(
          node.key.replace(/^[^:]+:/, "spill-meta:")
        );

        node.addDependency(spillMetaNode);
        node.setSpillMetaNode(spillMetaNode);
        spillMetaNode.setEvaluationResult(evaluation);
        
        if (evaluation.type === "spilled-values") {
          if (node.cellAddress.colIndex === 5 && node.cellAddress.rowIndex === 0) {
            console.log(`  About to evaluate origin for F1`);
            console.log(`    Spill source: ${evaluation.source}`);
          }
          
          const originResult = evaluation.evaluate({ x: 0, y: 0 }, ctx);
          
          if (node.cellAddress.colIndex === 5 && node.cellAddress.rowIndex === 0) {
            console.log(`  Origin result for F1:`, originResult);
            if (originResult.type === "value") {
              console.log(`    Value:`, originResult.result);
            } else if (originResult.type === "awaiting-evaluation") {
              console.log(`    Still awaiting evaluation! Waiting for:`, originResult.waitingFor.key);
            }
          }
          
          node.setEvaluationResult(originResult);
        } else {
          // Spill failed - set error on origin cell
          node.setEvaluationResult(evaluation);
        }
      }
      return;
    }

    if (evaluation.type === "value") {
      if (node instanceof SpillMetaNode) {
        node.setEvaluationResult({
          type: "does-not-spill",
        });
        return;
      } else {
        node.setEvaluationResult(evaluation);
        return;
      }
    }

    node.setEvaluationResult(evaluation);
  }

  evaluateDependencyNode(dependencyKey: string): void {
    if (dependencyKey.startsWith("empty:")) {
      this.evaluateEmptyCell(
        this.dependencyManager.getEmptyCellNode(dependencyKey)
      );
      return;
    }
    if (dependencyKey.startsWith("range:")) {
      this.evaluateRangeNode(
        this.dependencyManager.getRangeNode(dependencyKey)
      );
      return;
    }
    if (dependencyKey.startsWith("cell-value:")) {
      const node =
        this.dependencyManager.getCellValueOrEmptyCellNode(dependencyKey);
      if (node instanceof EmptyCellEvaluationNode) {
        this.evaluateEmptyCell(node);
        return;
      }
      this.evaluateCellNode(node);
      return;
    }
    if (dependencyKey.startsWith("ast:")) {
      // we could later move the evaluation logic here,
      // but for now, let's not do anything and
      // let the evaluateCellNode handle the evaluation
      // through formulaEvaluator.evaluateNode
      return;
    }
    if (dependencyKey.startsWith("spill-meta:")) {
      const node =
        this.dependencyManager.getSpillMetaOrEmptySpillMetaNode(dependencyKey);
      if (node instanceof EmptyCellEvaluationNode) {
        this.evaluateEmptyCell(node);
        return;
      }
      this.evaluateCellNode(node);
      return;
    }
    throw new Error("Invalid dependency key: " + dependencyKey);
  }

  /**
   * Evaluates a cell by building the evaluation order and evaluating the dependencies in order
   */
  evaluateCell(node: CellValueNode | EmptyCellEvaluationNode): void {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }
    this.isEvaluating = true;
    const sheet = this.workbookManager.getSheet(node.cellAddress);
    if (!sheet) {
      this.isEvaluating = false;
      throw new SheetNotFoundError(node.cellAddress.sheetName);
    }

    let precalculatedPlan: EvaluationOrder | undefined;

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;

      // Use DependencyManager to build evaluation order
      const evaluationPlan =
        precalculatedPlan ?? this.dependencyManager.buildEvaluationOrder(node);

      if (evaluationPlan.hasCycle) {
        const evaluationResult: ErrorEvaluationResult = {
          type: "error",
          err: FormulaError.CYCLE,
          message: Array.from(evaluationPlan.cycleNodes ?? [])
            .map((node) => node.key)
            .join(" -> "),
          errAddress: node,
        };
        // cycle detected
        if (evaluationPlan.cycleNodes) {
          for (const node of evaluationPlan.cycleNodes) {
            if (!(node instanceof RangeEvaluationNode)) {
              if (node instanceof AstEvaluationNode) {
                node.setEvaluationResult(evaluationResult);
              } else {
                node.setEvaluationResult(evaluationResult);
              }
            }
          }
        }
        this.isEvaluating = false;
      }

      // Evaluate all dependencies in order
      const timeStart = performance.now();
      const durations: { duration: number; key: string }[] = [];
      let numResolved = 0;
      if (flags.isProfiling && evaluationPlan.evaluationOrder.size > 2000) {
        // console.profile();
      }
      evaluationPlan.evaluationOrder.forEach((dependency) => {
        const start = performance.now();
        if (dependency.resolved) {
          numResolved++;
        }
        this.evaluateDependencyNode(dependency.key);
        const end = performance.now();
        if (flags.isProfiling && evaluationPlan.evaluationOrder.size > 2000) {
          durations.push({ duration: end - start, key: dependency.key });
        }
      });
      if (flags.isProfiling && evaluationPlan.evaluationOrder.size > 2000) {
        // console.profileEnd();
      }
      if (flags.isProfiling && evaluationPlan.evaluationOrder.size > 2000) {
        const percentResolved = Math.round(
          (100 * numResolved) / evaluationPlan.evaluationOrder.size
        );
        const avgDuration =
          durations.reduce((a, b) => a + b.duration, 0) / durations.length || 0;
        console.log(
          `%c[Evaluation] %c${evaluationPlan.evaluationOrder.size} deps | %c${(performance.now() - timeStart).toFixed(1)}ms | %c${percentResolved}% resolved | %c${avgDuration.toFixed(2)}ms avg`,
          "color:#83aaff;font-weight:bold;",
          "color:#fff;font-weight:bold;",
          "color:#7fff9e",
          "color:#85baff",
          "color:#ffdfa3"
        );
      }

      // let's check which nodes can be considered resolved
      this.dependencyManager.markResolvedNodes(node);

      const nextEvaluationPlan =
        this.dependencyManager.buildEvaluationOrder(node);

      precalculatedPlan = nextEvaluationPlan;

      // Check if new dependencies were discovered during evaluation
      if (nextEvaluationPlan.hash !== evaluationPlan.hash) {
        requiresReRun = true;
      } else {
        this.isEvaluating = false;
        return;
      }
    }
    this.isEvaluating = false;
  }

  convertScalarValueToCellValue(val: SerializedCellValue): CellValue {
    if (typeof val === "number") {
      return { type: "number", value: val };
    }
    if (typeof val === "boolean") {
      return { type: "boolean", value: val };
    }
    if (typeof val === "undefined") {
      return { type: "string", value: "" };
    }
    return { type: "string", value: val };
  }

  // todo optimize using workbook manager
  canSpill(spillCandidate: CellAddress, spillArea: SpreadsheetRange): boolean {
    const sheet = this.workbookManager.getSheet(spillCandidate);
    if (!sheet) {
      throw new SheetNotFoundError(spillCandidate.sheetName);
    }
    const cellId = getCellReference(spillCandidate);
    const content = sheet.content.get(cellId);
    if (!content) {
      throw new EvaluationError(FormulaError.REF, `Cell not found: ${cellId}`);
    }
    for (const spilledValue of this.dependencyManager.spilledValues) {
      if (
        spilledValue.origin.workbookName !== spillCandidate.workbookName ||
        spilledValue.origin.sheetName !== spillCandidate.sheetName
      ) {
        continue;
      }
      if (
        spilledValue.origin.colIndex === spillCandidate.colIndex &&
        spilledValue.origin.rowIndex === spillCandidate.rowIndex
      ) {
        // we are already have a spill, this one will be replaced
        continue;
      }

      if (checkRangeIntersection(spillArea, spilledValue.spillOnto)) {
        return false;
      }
    }
    // let's just check the raw data if there is something in the range
    for (const key of sheet.content.keys()) {
      const cellAddress = parseCellReference(key);
      const endCol = spillArea.end.col;
      const endRow = spillArea.end.row;

      if (
        cellAddress.colIndex === spillCandidate.colIndex &&
        cellAddress.rowIndex === spillCandidate.rowIndex
      ) {
        continue;
      }

      if (endCol.type === "number" && endRow.type === "number") {
        if (
          cellAddress.colIndex >= spillArea.start.col &&
          cellAddress.colIndex <= endCol.value &&
          cellAddress.rowIndex >= spillArea.start.row &&
          cellAddress.rowIndex <= endRow.value
        ) {
          if (
            normalizeSerializedCellValue(sheet.content.get(key)) !== undefined
          ) {
            // there is something in the range, so we can't spill
            return false;
          }
        }
      }
    }

    return true;
  }

  getCellEvaluationResult(
    cellAddress: CellAddress
  ): SingleEvaluationResult | undefined {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }

    const nodeKey = cellAddressToKey(cellAddress);
    const node = this.dependencyManager.getCellValueOrEmptyCellNode(nodeKey);

    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      throw new SheetNotFoundError(cellAddress.sheetName);
    }

    if (node.evaluationResult.type === "awaiting-evaluation") {
      if (cellAddressToKey(cellAddress).includes("F10")) {
        console.group("Evaluation of F10");
        flags.isProfiling = true;
        console.time("Evaluation of F10");
        // console.profile("Evaluation of F10");
      }
      this.evaluateCell(node);
      if (flags.isProfiling) {
        flags.isProfiling = false;
        console.timeEnd("Evaluation of F10");
        // console.profileEnd("Evaluation of F10");
        console.groupEnd();
      }
    }

    const result = node.evaluationResult;

    return result;
  }
}
