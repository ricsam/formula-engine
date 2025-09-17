import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type ErrorEvaluationResult,
  type EvaluatedDependencyNode,
  type EvaluationContext,
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
import { checkRangeIntersection } from "src/functions/math/open-range-evaluator";

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private formulaEvaluator: FormulaEvaluator,
    private storeManager: StoreManager
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

  getNodeDeps(nodeKey: string): Set<string> {
    const deps = new Set<string>();
    const node = this.storeManager.getEvaluatedNode(nodeKey);
    node?.deps?.forEach((dep) => deps.add(dep));
    node?.frontierDependencies?.forEach((frontierDep) => {
      if (node?.discardedFrontierDependencies?.has(frontierDep)) {
        return;
      }
      return deps.add(frontierDep);
    });
    return deps;
  }

  getTransitiveDeps(
    /**
     * nodeKey is the dependency node key, from dependencyNodeToKey
     */
    nodeKey: string
  ): Set<string> {
    const visited = new Set<string>();
    const queue = [nodeKey];

    let selfReferenced = false;

    while (queue.length > 0) {
      const current = queue.shift()!;

      if (visited.has(current)) {
        if (current === nodeKey) {
          selfReferenced = true;
        }
        continue;
      }
      visited.add(current);

      const deps = this.getNodeDeps(current);

      for (const dep of deps) {
        queue.push(dep);
      }
    }
    if (!selfReferenced) {
      visited.delete(nodeKey); // Don't include the starting node in the result
    }
    return visited;
  }

  private topologicalSort(
    /**
     * nodeKeys is the set of dependency node keys, see dependencyNodeToKey
     */
    nodeKeys: Set<string>
  ):
    | {
        type: "success";
        sorted: string[];
      }
    | {
        type: "cycle";
        inCycle: string[];
      } {
    const inDegree = new Map<string, number>();
    const queue: string[] = [];
    const result: string[] = [];

    // Calculate in-degrees
    for (const node of nodeKeys) {
      inDegree.set(node, 0);
    }

    for (const node of nodeKeys) {
      const deps = this.getNodeDeps(node);
      for (const precedent of deps) {
        inDegree.set(precedent, (inDegree.get(precedent) || 0) + 1);
      }
    }

    // Find nodes with no incoming edges
    for (const [node, degree] of inDegree) {
      if (degree === 0) {
        queue.push(node);
      }
    }

    // Process queue
    while (queue.length > 0) {
      const current = queue.shift()!;
      result.push(current);

      const deps = this.getNodeDeps(current);

      for (const dependent of deps) {
        const degree = inDegree.get(dependent)! - 1;
        inDegree.set(dependent, degree);

        if (degree === 0) {
          queue.push(dependent);
        }
      }
    }

    // Check if all nodes were processed (no cycles)
    if (result.length !== nodeKeys.size) {
      // Find nodes that are part of the cycle
      // These are nodes that still have non-zero in-degree
      const inCycle: string[] = [];
      for (const [node, degree] of inDegree) {
        if (degree > 0) {
          inCycle.push(node);
        }
      }

      return {
        type: "cycle",
        inCycle,
      };
    }

    return {
      type: "success",
      sorted: result,
    };
  }

  evaluateSpilled(
    cellAddress: CellAddress,
    context: EvaluationContext
  ):
    | { isSpilled: true; result: FunctionEvaluationResult | undefined }
    | { isSpilled: false } {
    const spilled = this.storeManager.getSpillValue(cellAddress);
    if (spilled) {
      const spillSource = this.storeManager.getSpilledAddress(
        cellAddress,
        spilled
      );
      const spillOrigin = this.storeManager.evalTimeSafeEvaluateCell(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return {
          isSpilled: true,
          result: spillOrigin.evaluate(spillSource.spillOffset, context),
        };
      }
    }
    return { isSpilled: false };
  }

  evaluateDependencyNode(
    /**
     * nodeKey is the dependency node key, from dependencyNodeToKey
     */
    nodeKey: string
  ): ValueEvaluationResult | ErrorEvaluationResult {
    const node = keyToDependencyNode(nodeKey);

    const dependenciesDiscoveredInEvaluation: Set<string> = new Set();
    const frontierDependenciesDiscoveredInEvaluation: Set<string> = new Set();
    const discardedFrontierDependenciesDiscoveredInEvaluation: Set<string> =
      new Set();

    let evaluation: FunctionEvaluationResult | undefined;

    const cellId = getCellReference({
      rowIndex: node.address.rowIndex,
      colIndex: node.address.colIndex,
    });

    const nodeAddress: CellAddress = {
      workbookName: node.workbookName,
      sheetName: node.sheetName,
      colIndex: node.address.colIndex,
      rowIndex: node.address.rowIndex,
    };

    const sheet = this.workbookManager.getSheet(nodeAddress);

    if (!sheet) {
      const depNode = {
        evaluationResult: {
          type: "error",
          err: FormulaError.REF,
          message: "Sheet not found",
        },
      } satisfies EvaluatedDependencyNode;
      this.storeManager.setEvaluatedNode(nodeKey, depNode);
      return depNode.evaluationResult;
    }

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
      return depNode.evaluationResult;
    }

    const evaluationContext: EvaluationContext = {
      currentCell: nodeAddress,
      dependencies: dependenciesDiscoveredInEvaluation,
      frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
      discardedFrontierDependencies:
        discardedFrontierDependenciesDiscoveredInEvaluation,
    };

    if (typeof content !== "string" || !content.startsWith("=")) {
      if (content !== "") {
        const depNode = {
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        } satisfies EvaluatedDependencyNode;
        this.storeManager.setEvaluatedNode(nodeKey, depNode);
        return depNode.evaluationResult;
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
          evaluationContext
        );
        if (spillOrigin && spillOrigin.type === "spilled-values") {
          // let's evaluate the spilled value to extract dependencies
          evaluation = spillOrigin.evaluate(
            spillTarget.spillOffset,
            evaluationContext
          );
        }
      } else {
        const currentDepNode = this.storeManager.getEvaluatedNode(nodeKey);
        const frontierDependencies: Set<string> =
          currentDepNode?.frontierDependencies ??
          new Set(
            this.workbookManager
              .getFrontierCandidates(
                {
                  start: {
                    col: nodeAddress.colIndex,
                    row: nodeAddress.rowIndex,
                  },
                  end: {
                    col: { type: "number", value: nodeAddress.colIndex },
                    row: { type: "number", value: nodeAddress.rowIndex },
                  },
                },
                nodeAddress
              )
              .map((candidate) =>
                dependencyNodeToKey({
                  address: candidate,
                  sheetName: candidate.sheetName,
                  workbookName: candidate.workbookName,
                })
              )
          );

        const discardedFrontierDependencies: Set<string> =
          currentDepNode?.discardedFrontierDependencies ?? new Set();

        const depNode = {
          frontierDependencies,
          discardedFrontierDependencies,
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        } satisfies EvaluatedDependencyNode;
        this.storeManager.setEvaluatedNode(nodeKey, depNode);
        return depNode.evaluationResult;
      }
    } else {
      evaluation = this.formulaEvaluator.evaluateFormula(
        content.slice(1),
        evaluationContext
      );
    }
    // below we are operating on a formula evaluation result

    //#region Check if the dependencies have changed
    const currentDeps =
      this.storeManager.getEvaluatedNode(nodeKey)?.deps ?? new Set();
    const currentFrontierDeps =
      this.storeManager.getEvaluatedNode(nodeKey)?.frontierDependencies ??
      new Set();

    let foundDeps = false;
    if (
      !(
        dependenciesDiscoveredInEvaluation.isSubsetOf(currentDeps) &&
        currentDeps.isSubsetOf(dependenciesDiscoveredInEvaluation)
      ) ||
      currentDeps.size !== dependenciesDiscoveredInEvaluation.size
    ) {
      foundDeps = true;
    }
    if (
      !(
        frontierDependenciesDiscoveredInEvaluation.isSubsetOf(
          currentFrontierDeps
        ) &&
        currentFrontierDeps.isSubsetOf(
          frontierDependenciesDiscoveredInEvaluation
        )
      ) ||
      currentFrontierDeps.size !==
        frontierDependenciesDiscoveredInEvaluation.size
    ) {
      foundDeps = true;
    }
    //#endregion

    // if this is the final evaluation, and it didn't yield any spill, then
    // it can not be a true frontier dependency
    if (!foundDeps) {
      this.storeManager.discardFrontierCandidate(nodeKey);
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

          // ooh shoot, we are spilling! Maybe some cells where we spill onto
          // previously discarded us as a frontier dependency because we returned an error
          // let's just remove us from the discarded frontier dependencies
          // and they will become re-evaluated because most likely I am now a parent
          // and evaluated before them in the loop of the sorted transient dependencies
          this.storeManager.restoreFrontierCandidate(nodeKey, spillArea);
        } else {
          evaluation = {
            type: "error",
            err: FormulaError.SPILL,
            message: "Can't spill",
          };
        }
      }
    }

    const depNode = {
      deps: dependenciesDiscoveredInEvaluation,
      frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
      discardedFrontierDependencies:
        discardedFrontierDependenciesDiscoveredInEvaluation,
      evaluationResult: evaluation,
    } satisfies EvaluatedDependencyNode;

    this.storeManager.setEvaluatedNode(nodeKey, depNode);

    let returnResult: ValueEvaluationResult | ErrorEvaluationResult | undefined;
    if (evaluation) {
      if (evaluation.type !== "spilled-values") {
        returnResult = evaluation;
      } else {
        // evaluation.
        const originEvaluation = evaluation.evaluate(
          { x: 0, y: 0 },
          evaluationContext
        );
        if (originEvaluation) {
          returnResult = originEvaluation;
        }
      }
    }

    return (
      returnResult ?? {
        type: "error",
        err: FormulaError.ERROR,
        message: "Evaluation failed",
      }
    );
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

      const allDeps = this.getTransitiveDeps(nodeKey);
      const sortResult = this.topologicalSort(allDeps);

      if (sortResult.type === "cycle") {
        const getDepNode = (nodeKey: string) =>
          ({
            deps:
              this.storeManager.getEvaluatedNode(nodeKey)?.deps ?? new Set(),
            frontierDependencies:
              this.storeManager.getEvaluatedNode(nodeKey)
                ?.frontierDependencies ?? new Set(),
            discardedFrontierDependencies:
              this.storeManager.getEvaluatedNode(nodeKey)
                ?.discardedFrontierDependencies ?? new Set(),
            evaluationResult: {
              type: "error",
              err: FormulaError.CYCLE,
              message: "Cycle detected",
            },
          }) satisfies EvaluatedDependencyNode;

        // cycle detected
        for (const nodeKey of sortResult.inCycle) {
          const depNode = getDepNode(nodeKey);
          this.storeManager.setEvaluatedNode(nodeKey, depNode);
        }
        this.isEvaluating = false;
        const depNode = getDepNode(nodeKey);
        this.storeManager.setEvaluatedNode(nodeKey, depNode);
        return depNode.evaluationResult;
      }

      const sorted = sortResult.sorted.reverse();
      sorted.forEach((nodeKey) => this.evaluateDependencyNode(nodeKey));
      const cellResult = this.evaluateDependencyNode(nodeKey);

      const transitiveDeps2 = this.getTransitiveDeps(nodeKey);

      if (
        allDeps.size !== transitiveDeps2.size ||
        !allDeps.isSubsetOf(transitiveDeps2)
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
    for (const value of this.storeManager.spilledValues.values()) {
      if (isCellInRange(originCellAddress, value.spillOnto)) {
        if (
          value.origin.colIndex === originCellAddress.colIndex &&
          value.origin.rowIndex === originCellAddress.rowIndex
        ) {
          continue;
        }
        return false;
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

    if (!value) {
      this.evaluateCell(cellAddress);
      value = getEvaluatedNode();
    }

    if (!value || !value.evaluationResult) {
      // nothing in the cell
      return undefined;
    }

    const result = value.evaluationResult;
    if (result.type === "spilled-values") {
      // the spill origin. Let's evaluate it
      // the origin should have been evaluated before and thus be
      // part of the dependency graph,
      // so we can just evaluate it here with a dummy context
      const dummyContext: EvaluationContext = {
        currentCell: cellAddress,
        dependencies: new Set(),
        frontierDependencies: new Set(),
        discardedFrontierDependencies: new Set(),
      };
      return result.evaluate({ x: 0, y: 0 }, dummyContext);
    }
    return result;
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.formulaEvaluator.isCellInTable(cellAddress);
  }
}
