import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type ErrorEvaluationResult,
  type EvaluatedDependencyNode,
  type FunctionEvaluationResult,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpilledValue,
  type SpreadsheetRange,
  type TableDefinition,
  type ValueEvaluationResult,
} from "../types";
import {
  getCellReference,
  isCellInRange,
  isRangeOneCell,
  parseCellReference,
} from "../utils";
import {
  dependencyNodeToKey,
  keyToDependencyNode,
} from "../utils/dependency-node-key";
import type { StoreManager } from "./store-manager";
import type { WorkbookManager } from "./workbook-manager";
import type { DependencyManager } from "./dependency-manager";
import { checkRangeIntersection } from "src/evaluator/open-range-evaluator";
import { EvaluationContext } from "src/evaluator/evaluation-context";

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private formulaEvaluator: FormulaEvaluator,
    private storeManager: StoreManager,
    private dependencyManager: DependencyManager
  ) {}

  getEvaluatedNodes() {
    return this.storeManager.getEvaluatedNodes();
  }

  getSpilledValues(): Map<string, SpilledValue> {
    return this.storeManager.spilledValues;
  }

  clearEvaluationCache(): void {
    this.storeManager.clearEvaluationCache();
  }

  evaluationResultToSerializedValue(
    evaluation: SingleEvaluationResult,
    debug?: boolean
  ): SerializedCellValue {
    if (evaluation.type !== "error") {
      const value = evaluation.result;

      return value.type === "infinity"
        ? value.sign === "positive"
          ? "INFINITY"
          : "-INFINITY"
        : value.value;
    }

    if (debug) {
      return evaluation.err + ": " + evaluation.message;
    }

    return evaluation.err;
  }

  evaluateDependencyNode(
    /**
     * nodeKey is the dependency node key, from dependencyNodeToKey
     */
    nodeKey: string
  ): SingleEvaluationResult {
    const node = keyToDependencyNode(nodeKey);
    const nodeAddress: CellAddress = {
      workbookName: node.workbookName,
      sheetName: node.sheetName,
      colIndex: node.address.colIndex,
      rowIndex: node.address.rowIndex,
    };
    const cellId = getCellReference({
      rowIndex: node.address.rowIndex,
      colIndex: node.address.colIndex,
    });

    const currentDepNode = this.storeManager.getEvaluatedNode(nodeKey);

    // Enable caching for resolved nodes
    if (currentDepNode?.resolved) {
      const result = currentDepNode.evaluationResult;
      if (result && result.type === "value") {
        return result;
      }
    }

    const ctx = new EvaluationContext(
      this.dependencyManager,
      this.storeManager,
      nodeAddress,
      currentDepNode
    );

    let evaluation: FunctionEvaluationResult | undefined;

    const sheet = this.workbookManager.getSheet(nodeAddress);

    if (!sheet) {
      const evaluationResult: ErrorEvaluationResult = {
        type: "error",
        err: FormulaError.REF,
        message: "Sheet not found",
      };
      this.storeManager.setEvaluatedNode(
        nodeKey,
        ctx.getEvaluatedDependencyNode(evaluationResult)
      );
      return evaluationResult;
    }

    let content: SerializedCellValue;
    try {
      content = normalizeSerializedCellValue(sheet.content.get(cellId));
    } catch (err) {
      const evaluationResult: ErrorEvaluationResult = {
        type: "error",
        err: FormulaError.ERROR,
        message: "Syntax error",
      };
      this.storeManager.setEvaluatedNode(
        nodeKey,
        ctx.getEvaluatedDependencyNode(evaluationResult)
      );
      return evaluationResult;
    }

    if (typeof content !== "string" || !content.startsWith("=")) {
      if (content !== "") {
        // Static value cells cannot have frontier dependencies
        const result: ValueEvaluationResult = {
          type: "value",
          result: this.convertScalarValueToCellValue(content),
        };
        this.storeManager.setEvaluatedNode(
          nodeKey,
          ctx.getEvaluatedDependencyNode(result)
        );
        return result;
      }
      // content === "", it is an empty cell, check if it has a spilled value
      const spilled = this.storeManager.getSpillValue(nodeAddress);
      if (spilled) {
        const spillTarget = this.storeManager.getSpilledAddress(
          nodeAddress,
          spilled
        );
        const spillOrigin = this.storeManager.evalTimeSafeEvaluateCell(
          spilled.origin,
          ctx
        );
        if (spillOrigin && spillOrigin.type === "spilled-values") {
          // let's evaluate the spilled value to extract dependencies
          evaluation = spillOrigin.evaluate(spillTarget.spillOffset, ctx);
        }
      } else {
        const emptyCellRange: SpreadsheetRange = {
          start: {
            col: nodeAddress.colIndex,
            row: nodeAddress.rowIndex,
          },
          end: {
            col: { type: "number", value: nodeAddress.colIndex },
            row: { type: "number", value: nodeAddress.rowIndex },
          },
        };
        // todo can be optimized to not generate the frontier candidates every time
        // we can cache the frontier candidates for the current cell
        const frontierCandidates: CellAddress[] =
          this.workbookManager.getFrontierCandidates(
            emptyCellRange,
            nodeAddress
          );

        for (const candidate of frontierCandidates) {
          const key = dependencyNodeToKey({
            address: candidate,
            sheetName: candidate.sheetName,
            workbookName: candidate.workbookName,
          });

          if (ctx.isFrontierDependencyDiscarded(key, emptyCellRange)) {
            continue;
          }

          const node = this.storeManager.getEvaluatedNode(key);

          const result =
            this.storeManager.getEvaluatedNode(key)?.evaluationResult;

          ctx.addFrontierDependency(key, emptyCellRange);

          // upgrade or downgrade frontier dependency
          if (result) {
            if (result.type === "spilled-values") {
              const spillArea = result.spillArea(candidate);
              const intersects = isCellInRange(nodeAddress, spillArea);
              if (intersects) {
                ctx.maybeUpgradeFrontierDependency(key, emptyCellRange); // upgraded!
              } else {
                ctx.maybeDiscardFrontierDependency(key, emptyCellRange); // downgraded!
              }
            } else {
              ctx.maybeDiscardFrontierDependency(key, emptyCellRange); // downgraded!
            }
          }
        }

        const evaluationResult: SingleEvaluationResult = {
          type: "value",
          result: this.convertScalarValueToCellValue(content),
        };
        this.storeManager.setEvaluatedNode(
          nodeKey,
          ctx.getEvaluatedDependencyNode(evaluationResult)
        );
        return evaluationResult;
      }
    } else {
      evaluation = this.formulaEvaluator.evaluateFormula(content.slice(1), ctx);
    }

    // if a cell returns a range, we need to spill the values onto the sheet
    if (evaluation && evaluation.type === "spilled-values") {
      const spillArea = evaluation.spillArea(nodeAddress);
      if (!isRangeOneCell(spillArea)) {
        if (this.canSpill(nodeAddress, spillArea)) {
          this.storeManager.spilledValues.set(nodeKey, {
            spillOnto: spillArea,
            origin: nodeAddress,
          });
        } else {
          evaluation = {
            type: "error",
            err: FormulaError.SPILL,
            message: "Can't spill",
          };
        }
      } else {
        throw new Error("We should not be able to spill a single cell");
      }
    }

    let returnResult: SingleEvaluationResult | undefined;
    let originSpillResult: SingleEvaluationResult | undefined;
    if (evaluation) {
      if (evaluation.type !== "spilled-values") {
        returnResult = evaluation;
      } else {
        // for the spilled origin we need to evaluate the origin and return the result
        originSpillResult = evaluation.evaluate({ x: 0, y: 0 }, ctx);
        if (originSpillResult) {
          returnResult = originSpillResult;
        }
      }
    }

    const failedEvaluation: ErrorEvaluationResult = {
      type: "error",
      err: FormulaError.ERROR,
      message: "Evaluation failed",
    };

    this.storeManager.setEvaluatedNode(
      nodeKey,
      // we store the evaluation result from evaluateFormula
      ctx.getEvaluatedDependencyNode(
        evaluation ?? failedEvaluation,
        originSpillResult
      )
    );

    return returnResult ?? failedEvaluation;
  }

  evaluateCell(
    cellAddress: CellAddress,
    cycleCheck: boolean = false
  ): ValueEvaluationResult | ErrorEvaluationResult {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }
    this.isEvaluating = true;
    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      this.isEvaluating = false;
      throw new Error("Sheet not found");
    }

    const cellId = getCellReference({
      rowIndex: cellAddress.rowIndex,
      colIndex: cellAddress.colIndex,
    });

    const nodeKey = dependencyNodeToKey({
      address: cellAddress,
      sheetName: cellAddress.sheetName,
      workbookName: cellAddress.workbookName,
    });

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;
      let content: SerializedCellValue;
      try {
        content = normalizeSerializedCellValue(sheet.content.get(cellId));
      } catch (err) {
        const depNode = {
          evaluationResult: {
            type: "error",
            err: FormulaError.ERROR,
            message: "Syntax error",
          },
        } satisfies EvaluatedDependencyNode;

        this.storeManager.setEvaluatedNode(nodeKey, depNode);
        this.isEvaluating = false;
        return depNode.evaluationResult;
      }
      if (
        typeof content !== "string" ||
        // "" may be spilled, so should be evaluated
        (!content.startsWith("=") && content !== "")
      ) {
        const depNode = {
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        } satisfies EvaluatedDependencyNode;
        this.storeManager.setEvaluatedNode(nodeKey, depNode);
        this.isEvaluating = false;
        return depNode.evaluationResult;
      }

      // Use DependencyManager to build evaluation order
      const evaluationPlan =
        this.dependencyManager.buildEvaluationOrder(nodeKey);

      if (evaluationPlan.hasCycle) {
        // TODO: fix this
        const getDepNode = (nodeKey: string) =>
          ({
            deps:
              this.storeManager.getEvaluatedNode(nodeKey)?.deps ?? new Set(),
            frontierDependencies:
              this.storeManager.getEvaluatedNode(nodeKey)
                ?.frontierDependencies ?? new Map(),
            discardedFrontierDependencies:
              this.storeManager.getEvaluatedNode(nodeKey)
                ?.discardedFrontierDependencies ?? new Map(),
            evaluationResult: {
              type: "error",
              err: FormulaError.CYCLE,
              message: "Cycle detected",
            },
          }) satisfies EvaluatedDependencyNode;

        // cycle detected
        if (evaluationPlan.cycleNodes) {
          for (const nodeKey of evaluationPlan.cycleNodes) {
            const depNode = getDepNode(nodeKey);
            this.storeManager.setEvaluatedNode(nodeKey, depNode);
          }
        }
        this.isEvaluating = false;
        const depNode = getDepNode(nodeKey);
        this.storeManager.setEvaluatedNode(nodeKey, depNode);
        return depNode.evaluationResult;
      }

      // Evaluate all dependencies in order
      evaluationPlan.evaluationOrder.forEach((dependency) =>
        {
          if (dependency === nodeKey) {
            return;
          }
          return this.evaluateDependencyNode(dependency);
        }
      );
      const cellResult = this.evaluateDependencyNode(nodeKey);

      // Check if new dependencies were discovered during evaluation

      if (
        this.dependencyManager.buildEvaluationOrder(nodeKey).hash !==
        evaluationPlan.hash
      ) {
        requiresReRun = true;
      } else {
        // check for cycles once we have finished
        if (!cycleCheck) {
          this.isEvaluating = false;
          return this.evaluateCell(cellAddress, true);
        } else {
          this.isEvaluating = false;
          return cellResult;
        }
      }
    }
    this.isEvaluating = false;
    return {
      type: "error",
      err: FormulaError.ERROR,
      message: "Evaluation failed",
    };
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
  canSpill(originCellAddress: CellAddress, range: SpreadsheetRange): boolean {
    const sheet = this.workbookManager.getSheet(originCellAddress);
    if (!sheet) {
      throw new Error("Sheet not found");
    }
    const cellId = getCellReference(originCellAddress);
    const content = sheet.content.get(cellId);
    if (!content) {
      throw new Error(`Cell not found: ${cellId}`);
    }
    for (const spilledValue of this.storeManager.spilledValues.values()) {
      if (
        spilledValue.origin.workbookName === originCellAddress.workbookName &&
        spilledValue.origin.sheetName === originCellAddress.sheetName &&
        spilledValue.origin.colIndex === originCellAddress.colIndex &&
        spilledValue.origin.rowIndex === originCellAddress.rowIndex
      ) {
        continue;
      }
      if (checkRangeIntersection(range, spilledValue.spillOnto)) {
        return false;
      }
    }
    // let's just check the raw data if there is something in the range
    for (const key of sheet.content.keys()) {
      const cellAddress = parseCellReference(key);
      const endCol = range.end.col;
      const endRow = range.end.row;

      if (
        cellAddress.colIndex === originCellAddress.colIndex &&
        cellAddress.rowIndex === originCellAddress.rowIndex
      ) {
        continue;
      }

      if (endCol.type === "number" && endRow.type === "number") {
        if (
          cellAddress.colIndex >= range.start.col &&
          cellAddress.colIndex <= endCol.value &&
          cellAddress.rowIndex >= range.start.row &&
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

    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    const getEvaluatedNode = () => {
      return this.storeManager.getEvaluatedNode(
        dependencyNodeToKey({
          address: cellAddress,
          sheetName: cellAddress.sheetName,
          workbookName: cellAddress.workbookName,
        })
      );
    };

    let value = getEvaluatedNode();

    if (
      !value ||
      (value &&
        value.evaluationResult &&
        value.evaluationResult.type === "spilled-values" &&
        !value.originSpillResult)
    ) {
      this.evaluateCell(cellAddress);
      value = getEvaluatedNode();
    }

    if (!value || !value.evaluationResult) {
      // nothing in the cell
      return undefined;
    }

    const result = value.originSpillResult ?? value.evaluationResult;
    if (result.type === "spilled-values") {
      throw new Error("Spilled values should have been evaluated before");
    }
    return result;
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.formulaEvaluator.isCellInTable(cellAddress);
  }
}
