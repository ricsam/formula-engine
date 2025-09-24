import { EvaluationContext } from "src/evaluator/evaluation-context";
import { checkRangeIntersection } from "src/evaluator/open-range-evaluator";
import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type ErrorEvaluationResult,
  type FunctionEvaluationResult,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpilledValue,
  type SpreadsheetRange,
  type TableDefinition,
  type ValueEvaluationResult,
} from "../types";
import {
  cellAddressToKey,
  getCellReference,
  isCellInRange,
  isRangeOneCell,
  keyToCellAddress,
  parseCellReference,
} from "../utils";
import type { DependencyManager } from "./dependency-manager";
import type { WorkbookManager } from "./workbook-manager";

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private formulaEvaluator: FormulaEvaluator,
    private storeManager: DependencyManager
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

  evaluateDependencyNode(cellReference: string): SingleEvaluationResult {
    const nodeAddress: CellAddress = keyToCellAddress(cellReference);
    const cellId = getCellReference({
      rowIndex: nodeAddress.rowIndex,
      colIndex: nodeAddress.colIndex,
    });

    const currentDepNode = this.storeManager.getEvaluatedNode(cellReference);

    // Enable caching for resolved nodes
    if (currentDepNode?.resolved) {
      const result = currentDepNode.evaluationResult;
      if (result && result.type === "value") {
        return result;
      }
    }

    const ctx = new EvaluationContext(
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
        cellReference,
        ctx.getDependencyAttributes(),
        evaluationResult
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
        cellReference,
        ctx.getDependencyAttributes(),
        evaluationResult
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
          cellReference,
          ctx.getDependencyAttributes(),
          result
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
          const key = cellAddressToKey(candidate);

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
          cellReference,
          ctx.getDependencyAttributes(),
          evaluationResult
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
          this.storeManager.spilledValues.set(cellReference, {
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
      cellReference,
      // we store the evaluation result from evaluateFormula
      ctx.getDependencyAttributes(),
      evaluation ?? failedEvaluation,
      originSpillResult
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

    const nodeKey = cellAddressToKey(cellAddress);

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;
      let content: SerializedCellValue;
      try {
        content = normalizeSerializedCellValue(sheet.content.get(cellId));
      } catch (err) {
        const evaluationResult: ErrorEvaluationResult = {
          type: "error",
          err: FormulaError.ERROR,
          message: "Syntax error",
        };
        this.storeManager.setEvaluatedResult(nodeKey, evaluationResult);
        this.isEvaluating = false;
        return evaluationResult;
      }
      if (
        typeof content !== "string" ||
        // "" may be spilled, so should be evaluated
        (!content.startsWith("=") && content !== "")
      ) {
        const evaluationResult: ValueEvaluationResult = {
          type: "value",
          result: this.convertScalarValueToCellValue(content),
        };
        this.storeManager.setEvaluatedResult(nodeKey, evaluationResult);
        this.isEvaluating = false;
        return evaluationResult;
      }

      // Use DependencyManager to build evaluation order
      const evaluationPlan = this.storeManager.buildEvaluationOrder(nodeKey);

      if (evaluationPlan.hasCycle) {
        const evaluationResult: ErrorEvaluationResult = {
          type: "error",
          err: FormulaError.CYCLE,
          message: "Cycle detected",
        };
        // cycle detected
        if (evaluationPlan.cycleNodes) {
          for (const nodeKey of evaluationPlan.cycleNodes) {
            this.storeManager.setEvaluatedResult(nodeKey, evaluationResult);
          }
        }
        this.isEvaluating = false;
        this.storeManager.setEvaluatedResult(nodeKey, evaluationResult);
        return evaluationResult;
      }

      // Evaluate all dependencies in order
      evaluationPlan.evaluationOrder.forEach((dependency) => {
        if (dependency === nodeKey) {
          return;
        }
        return this.evaluateDependencyNode(dependency);
      });
      const cellResult = this.evaluateDependencyNode(nodeKey);

      // Check if new dependencies were discovered during evaluation

      if (
        this.storeManager.buildEvaluationOrder(nodeKey).hash !==
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
    console.count("canSpill");
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
      return this.storeManager.getEvaluatedNode(cellAddressToKey(cellAddress));
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
