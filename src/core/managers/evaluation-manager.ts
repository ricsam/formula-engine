import { EvaluationContext } from "src/evaluator/evaluation-context";
import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type ErrorEvaluationResult,
  type EvaluationOrder,
  type FunctionEvaluationResult,
  type RangeAddress,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpilledValue,
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
  rangeAddressToKey,
} from "../utils";
import type { DependencyManager } from "./dependency-manager";
import type { WorkbookManager } from "./workbook-manager";
import { flags } from "src/debug/flags";
import { CellEvalNode } from "src/evaluator/cell-eval-node";
import { EmptyCellEvaluationNode } from "src/evaluator/empty-cell-evaluation-node";
import { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private formulaEvaluator: FormulaEvaluator,
    private dependencyManager: DependencyManager
  ) {}

  getSpilledValues(): Map<string, SpilledValue> {
    return this.dependencyManager.spilledValues;
  }

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
        cellAddressToKey(evaluation.errAddress) +
        " is awaiting evaluation of " +
        cellAddressToKey(evaluation.waitingFor)
      );
    }

    if (debug) {
      const errAddress = cellAddressToKey(evaluation.errAddress);
      if (errAddress === cellAddressToKey(cellAddress)) {
        return evaluation.err + " " + evaluation.message;
      }
      return (
        evaluation.err +
        " in " +
        cellAddressToKey(evaluation.errAddress) +
        " " +
        evaluation.message
      );
    }

    return evaluation.err;
  }

  evaluateEmptyCell(cellReference: string): void {
    const nodeAddress: CellAddress = keyToCellAddress(cellReference);

    const node = this.dependencyManager.getEmptyCellNode(cellReference);
    node.resetDirectDepsUpdated();

    if (node.resolved) {
      const result = node.evaluationResult;
      if (result && result.type !== "awaiting-evaluation") {
        return;
      }
    }

    const ctx = new EvaluationContext(node, node);
    const inSpilled = this.dependencyManager.getSpillValue(nodeAddress);

    if (inSpilled) {
      const spillTarget = this.dependencyManager.getSpilledAddress(
        nodeAddress,
        inSpilled
      );
      const spillOriginKey = cellAddressToKey(inSpilled.origin);
      const spillOrigin = this.dependencyManager.getCellNode(spillOriginKey);
      node.addDependency(spillOrigin);
      const result = spillOrigin.evaluationResult;
      if (result && result.type === "spilled-values") {
        // let's evaluate the spilled value to extract dependencies
        const evaluation = captureEvaluationErrors(spillTarget.address, () => {
          return result.evaluate(spillTarget.spillOffset, ctx);
        });
        node.setEvaluationResult(evaluation);
      }
    } else {
      const emptyCellRange: RangeAddress = {
        range: {
          start: {
            col: nodeAddress.colIndex,
            row: nodeAddress.rowIndex,
          },
          end: {
            col: { type: "number", value: nodeAddress.colIndex },
            row: { type: "number", value: nodeAddress.rowIndex },
          },
        },
        sheetName: nodeAddress.sheetName,
        workbookName: nodeAddress.workbookName,
      };
      // todo can be optimized to not generate the frontier candidates every time
      // we can cache the frontier candidates for the current cell
      const frontierCandidates: CellAddress[] =
        this.workbookManager.getFrontierCandidates(emptyCellRange);

      for (const candidate of frontierCandidates) {
        const key = cellAddressToKey(candidate);

        const candidateNode = this.dependencyManager.getCellNode(key);

        if (candidateNode instanceof EmptyCellEvaluationNode) {
          throw new Error("A frontier dependencies can not be an empty cell");
        }

        node.addFrontierDependency(candidateNode); // register the frontier dependency

        const result = candidateNode.evaluationResult;

        // upgrade or downgrade frontier dependency
        if (result) {
          if (result.type === "spilled-values") {
            const spillArea = result.spillArea(candidate);
            const intersects = isCellInRange(nodeAddress, spillArea);
            if (intersects) {
              node.maybeUpgradeFrontierDependency(candidateNode); // upgraded!
            } else {
              node.maybeDiscardFrontierDependency(candidateNode); // downgraded!
            }
          } else {
            node.maybeDiscardFrontierDependency(candidateNode); // downgraded!
          }
        }
      }

      const evaluationResult: SingleEvaluationResult = {
        type: "value",
        result: this.convertScalarValueToCellValue(""),
      };
      // for now let's just store the empty value, the next time the cell is evaluated isSpilled will be true and the spilled value will be evaluated
      node.setEvaluationResult(evaluationResult);
    }
  }

  evaluateRangeNode(dependencyKey: string): void {
    const node = this.dependencyManager.getRangeNode(dependencyKey);
    if (node.resolved) {
      return;
    }

    node.resetDirectDepsUpdated();
    // this is just about setting up the dependencies for the range node

    // let's setup the dependencies

    //#region frontier dependencies
    const frontierCandidates: CellAddress[] =
      this.workbookManager.getFrontierCandidates(node.address);

    for (const candidate of frontierCandidates) {
      const key = cellAddressToKey(candidate);

      const candidateNode = this.dependencyManager.getCellNode(key);

      if (candidateNode instanceof EmptyCellEvaluationNode) {
        throw new Error("A frontier dependencies can not be an empty cell");
      }

      node.addFrontierDependency(candidateNode); // register the frontier dependency

      const result = candidateNode.evaluationResult;

      // upgrade or downgrade frontier dependency
      if (result) {
        if (result.type === "spilled-values") {
          const spillArea = result.spillArea(candidate);
          const intersects = checkRangeIntersection(
            node.address.range,
            spillArea
          );
          if (intersects) {
            node.maybeUpgradeFrontierDependency(candidateNode); // upgraded!
          } else {
            node.maybeDiscardFrontierDependency(candidateNode); // downgraded!
          }
        } else {
          node.maybeDiscardFrontierDependency(candidateNode); // downgraded!
        }
      }
    }
    //#endregion

    //#region normal dependencies
    const cellsInRange = this.workbookManager.getCellsInRange(node.address);

    // Iterate over all defined cells in the sheet using optimized index-based iterator
    for (const address of cellsInRange) {
      const cellKey = cellAddressToKey(address);

      const cellNode = this.dependencyManager.getCellNode(cellKey);
      node.addDependency(cellNode);
    }
    //#endregion
  }

  evaluateCellNode(dependencyKey: string): void {
    const nodeAddress: CellAddress = keyToCellAddress(dependencyKey);
    const cellId = getCellReference({
      rowIndex: nodeAddress.rowIndex,
      colIndex: nodeAddress.colIndex,
    });

    const sheet = this.workbookManager.getSheet(nodeAddress);

    if (!sheet) {
      throw new Error("Sheet not found");
    }

    const rawContent = sheet.content.get(cellId);
    let isEmptyCell = false;
    if (rawContent === undefined) {
      isEmptyCell = true;
    } else if (typeof rawContent === "string") {
      if (rawContent === "") {
        isEmptyCell = true;
      }
    }

    if (isEmptyCell) {
      this.evaluateEmptyCell(dependencyKey);
      return;
    }

    const currentDepNode = this.dependencyManager.getCellNode(dependencyKey);

    if (currentDepNode instanceof EmptyCellEvaluationNode) {
      this.evaluateEmptyCell(dependencyKey);
      return;
    }

    // Enable caching for resolved nodes
    if (currentDepNode.resolved) {
      const result = currentDepNode.evaluationResult;
      if (result && result.type !== "awaiting-evaluation") {
        return;
      }
    }

    currentDepNode.resetDirectDepsUpdated();

    const ctx = new EvaluationContext(currentDepNode, currentDepNode);

    let content: SerializedCellValue;
    try {
      content = normalizeSerializedCellValue(rawContent);
    } catch (err) {
      const evaluationResult: ErrorEvaluationResult = {
        type: "error",
        err: FormulaError.ERROR,
        message: "Syntax error",
        errAddress: nodeAddress,
      };
      currentDepNode.setEvaluationResult(evaluationResult);
      return;
    }

    if (typeof content !== "string" || !content.startsWith("=")) {
      // Static value cells cannot have frontier dependencies
      const result: ValueEvaluationResult = {
        type: "value",
        result: this.convertScalarValueToCellValue(content),
      };
      currentDepNode.setEvaluationResult(result);
      return;
    }

    let evaluation: FunctionEvaluationResult =
      this.formulaEvaluator.evaluateFormula(content.slice(1), ctx);

    // if a cell returns a range, we need to spill the values onto the sheet
    if (evaluation.type === "spilled-values") {
      const spillArea = evaluation.spillArea(nodeAddress);
      if (!isRangeOneCell(spillArea)) {
        if (this.canSpill(nodeAddress, spillArea)) {
          this.dependencyManager.spilledValues.set(dependencyKey, {
            spillOnto: spillArea,
            origin: nodeAddress,
          });
        } else {
          evaluation = {
            type: "error",
            err: FormulaError.SPILL,
            message: "Can't spill",
            errAddress: nodeAddress,
          };
        }
      } else {
        throw new Error("We should not be able to spill a single cell");
      }
    }

    let originSpillResult: SingleEvaluationResult | undefined;
    if (evaluation) {
      if (evaluation.type === "spilled-values") {
        // for the spilled origin we need to evaluate the origin and store the result
        originSpillResult = captureEvaluationErrors(nodeAddress, () => {
          return evaluation.evaluate({ x: 0, y: 0 }, ctx);
        });
      }
    }

    const failedEvaluation: ErrorEvaluationResult = {
      type: "error",
      err: FormulaError.ERROR,
      message: "Evaluation failed",
      errAddress: nodeAddress,
    };

    currentDepNode.setEvaluationResult(
      // we store the evaluation result from evaluateFormula
      evaluation ?? failedEvaluation,
      originSpillResult
    );
  }

  evaluateDependencyNode(dependencyKey: string): void {
    if (dependencyKey.startsWith("empty:")) {
      this.evaluateEmptyCell(dependencyKey);
      return;
    }
    if (dependencyKey.startsWith("range:")) {
      this.evaluateRangeNode(dependencyKey);
      return;
    }
    if (dependencyKey.startsWith("cell:")) {
      this.evaluateCellNode(dependencyKey);
      return;
    }
    throw new Error("Invalid dependency key: " + dependencyKey);
  }

  /**
   * Evaluates a cell by building the evaluation order and evaluating the dependencies in order
   */
  evaluateCell(cellAddress: CellAddress): void {
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
    let precalculatedPlan: EvaluationOrder | undefined;

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;

      // Use DependencyManager to build evaluation order
      const evaluationPlan =
        precalculatedPlan ??
        this.dependencyManager.buildEvaluationOrder(nodeKey);

      if (evaluationPlan.hasCycle) {
        const evaluationResult: ErrorEvaluationResult = {
          type: "error",
          err: FormulaError.CYCLE,
          message: Array.from(evaluationPlan.cycleNodes ?? [])
            .map((node) => node.key)
            .join(" -> "),
          errAddress: cellAddress,
        };
        // cycle detected
        if (evaluationPlan.cycleNodes) {
          for (const node of evaluationPlan.cycleNodes) {
            if (!(node instanceof RangeEvaluationNode)) {
              node.setEvaluationResult({
                ...evaluationResult,
                errAddress: node.cellAddress,
              });
            }
          }
        }
        this.isEvaluating = false;
        return;
      }

      // Evaluate all dependencies in order
      evaluationPlan.evaluationOrder.forEach((dependency) => {
        this.evaluateDependencyNode(dependency.key);
      });

      const evalResult = this.dependencyManager.getCellNode(nodeKey);
      const failedEvaluation: ErrorEvaluationResult = {
        type: "error",
        err: FormulaError.ERROR,
        message: "Evaluation failed",
        errAddress: cellAddress,
      };

      /**
       * In reality, this is a SingleEvaluationResult
       */
      const cellResult: FunctionEvaluationResult =
        evalResult?.originSpillResult ??
        evalResult?.evaluationResult ??
        failedEvaluation;

      if (cellResult.type === "spilled-values") {
        throw new Error(
          "Spilled values should have been evaluated before, and originSpillResult should be set"
        );
      }

      // let's check which nodes can be considered resolved
      this.dependencyManager.markResolvedNodes(nodeKey);

      const nextEvaluationPlan =
        this.dependencyManager.buildEvaluationOrder(nodeKey);

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
      throw new Error("Sheet not found");
    }
    const cellId = getCellReference(spillCandidate);
    const content = sheet.content.get(cellId);
    if (!content) {
      throw new Error(`Cell not found: ${cellId}`);
    }
    for (const spilledValue of this.dependencyManager.spilledValues.values()) {
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

    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    const getEvaluatedNode = () => {
      return this.dependencyManager.getCellNode(cellAddressToKey(cellAddress));
    };

    let value = getEvaluatedNode();

    if (
      !value ||
      (value && !value.evaluationResult) ||
      (value && value.evaluationResult?.type === "awaiting-evaluation") ||
      (value.evaluationResult &&
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
