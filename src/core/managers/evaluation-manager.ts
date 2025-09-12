import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type EvaluationContext,
  type FunctionEvaluationResult,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpilledValue,
  type SpreadsheetRange,
  type TableDefinition,
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

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private formulaEvaluator: FormulaEvaluator,
    private storeManager: StoreManager
  ) {}

  getEvaluatedNodes() {
    return this.storeManager.evaluatedNodes;
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
    const node = this.storeManager.evaluatedNodes.get(nodeKey);
    node?.deps?.forEach((dep) => deps.add(dep));
    node?.frontierDependencies?.forEach((frontierDep) => {
      if (node?.discardedFrontierDependencies?.has(frontierDep)) {
        return;
      }
      return deps.add(frontierDep);
    });
    return deps;
  }

  /**
   * Discovers dependencies for a node without fully evaluating it.
   * This is used for cycle detection before evaluation.
   */
  private discoverNodeDeps(nodeKey: string): Set<string> {
    const node = keyToDependencyNode(nodeKey);
    const deps = new Set<string>();
    const cellAddress: CellAddress = {
      workbookName: node.workbookName,
      sheetName: node.sheetName,
      colIndex: node.address.colIndex,
      rowIndex: node.address.rowIndex,
    };

    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      return deps;
    }

    const cellId = getCellReference({
      rowIndex: node.address.rowIndex,
      colIndex: node.address.colIndex,
    });

    let content: SerializedCellValue;
    try {
      content = normalizeSerializedCellValue(sheet.content.get(cellId));
    } catch (err) {
      return deps;
    }

    if (typeof content === "string" && content.startsWith("=")) {
      // Parse the formula to discover dependencies without evaluating
      const dependenciesDiscoveredInEvaluation: Set<string> = new Set();
      const frontierDependenciesDiscoveredInEvaluation: Set<string> = new Set();
      const discardedFrontierDependenciesDiscoveredInEvaluation: Set<string> =
        new Set();

      const evaluationContext: EvaluationContext = {
        currentCell: cellAddress,
        dependencies: dependenciesDiscoveredInEvaluation,
        frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
        discardedFrontierDependencies:
          discardedFrontierDependenciesDiscoveredInEvaluation,
      };

      this.formulaEvaluator.evaluateFormula(
        content.slice(1),
        evaluationContext
      );

      dependenciesDiscoveredInEvaluation.forEach((dep) => deps.add(dep));
      frontierDependenciesDiscoveredInEvaluation.forEach((dep) =>
        deps.add(dep)
      );
    }

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

    while (queue.length > 0) {
      const current = queue.shift()!;

      if (visited.has(current)) continue;
      visited.add(current);

      const deps = this.getNodeDeps(current);

      for (const dep of deps) {
        queue.push(dep);
      }
    }

    visited.delete(nodeKey); // Don't include the starting node
    return visited;
  }

  /**
   * Discovers transitive dependencies without full evaluation.
   * This is used for cycle detection before evaluation.
   */
  private discoverTransitiveDeps(nodeKey: string): Set<string> {
    const visited = new Set<string>();
    const queue = [nodeKey];

    while (queue.length > 0) {
      const current = queue.shift()!;

      if (visited.has(current)) continue;
      visited.add(current);

      const deps = this.discoverNodeDeps(current);

      for (const dep of deps) {
        queue.push(dep);
      }
    }

    visited.delete(nodeKey); // Don't include the starting node
    return visited;
  }

  private topologicalSort(
    /**
     * nodeKeys is the set of dependency node keys, see dependencyNodeToKey
     */
    nodeKeys: Set<string>
  ): string[] | null {
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
      return null; // Cycle detected
    }

    return result;
  }

  /**
   * Finds all nodes that participate in dependency cycles within the given set of nodes.
   * Uses DFS to detect strongly connected components with more than one node or self-loops.
   */
  private findCycleParticipants(nodeKeys: Set<string>): Set<string> {
    const visited = new Set<string>();
    const recursionStack = new Set<string>();
    const cycleParticipants = new Set<string>();

    const dfs = (node: string): boolean => {
      if (recursionStack.has(node)) {
        // Found a back edge - this indicates a cycle
        // Mark all nodes in the current recursion stack as cycle participants
        cycleParticipants.add(node);
        return true;
      }

      if (visited.has(node)) {
        return false;
      }

      visited.add(node);
      recursionStack.add(node);

      const deps = this.discoverNodeDeps(node);
      let foundCycle = false;

      for (const dep of deps) {
        if (nodeKeys.has(dep)) {
          // Only consider dependencies within our node set
          if (dfs(dep)) {
            cycleParticipants.add(node);
            foundCycle = true;
          }
        }
      }

      recursionStack.delete(node);
      return foundCycle;
    };

    // Run DFS from each unvisited node
    for (const node of nodeKeys) {
      if (!visited.has(node)) {
        dfs(node);
      }
    }

    return cycleParticipants;
  }

  evaluateSpilled(
    cellAddress: CellAddress,
    context: EvaluationContext
  ):
    | { isSpilled: true; result: FunctionEvaluationResult | undefined }
    | { isSpilled: false } {
    const spilled = this.storeManager.isSpilled(cellAddress);
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
    nodeKey: string,
    /**
     * We evaluate the dependency node in the context of the cell address
     */
    cellAddress: CellAddress
  ): boolean {
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
      this.storeManager.evaluatedNodes.set(nodeKey, {
        evaluationResult: {
          type: "error",
          err: FormulaError.REF,
          message: "Sheet not found",
        },
      });
      return false;
    }

    let content: SerializedCellValue;
    try {
      content = normalizeSerializedCellValue(sheet.content.get(cellId));
    } catch (err) {
      this.storeManager.evaluatedNodes.set(nodeKey, {
        evaluationResult: {
          type: "error",
          err: FormulaError.ERROR,
          message: "Syntax error",
        },
      });
      return false;
    }

    const evaluationContext: EvaluationContext = {
      currentCell: nodeAddress,
      dependencies: dependenciesDiscoveredInEvaluation,
      frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
      discardedFrontierDependencies:
        discardedFrontierDependenciesDiscoveredInEvaluation,
    };

    if (typeof content !== "string" || !content.startsWith("=")) {
      const spilled = this.storeManager.isSpilled(nodeAddress);
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
        this.storeManager.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        });
        return false;
      }
    } else {
      evaluation = this.formulaEvaluator.evaluateFormula(
        content.slice(1),
        evaluationContext
      );
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

          const isDependencyInRange = (dep: string) => {
            const node = keyToDependencyNode(dep);
            const cellAddress: CellAddress = {
              ...node.address,
              sheetName: node.sheetName,
              workbookName: node.workbookName,
            };
            return isCellInRange(cellAddress, spillArea);
          };

          this.storeManager.evaluatedNodes.forEach((evaled, key) => {
            for (const dep of evaled.deps ?? []) {
              if (isDependencyInRange(dep)) {
                // one of the dependencies of key is in the spill area,
                // we need to re-evaluate the cell
                this.evaluateDependencyNode(key, cellAddress);
                return; // go to next evaluated node
              }
            }
          });
        } else {
          evaluation = {
            type: "error",
            err: FormulaError.SPILL,
            message: "Can't spill",
          };
        }
      }
    }

    const currentDeps =
      this.storeManager.evaluatedNodes.get(nodeKey)?.deps ?? new Set();
    const currentFrontierDeps =
      this.storeManager.evaluatedNodes.get(nodeKey)?.frontierDependencies ??
      new Set();

    let requiresReRun = true;
    if (
      !(
        dependenciesDiscoveredInEvaluation.isSubsetOf(currentDeps) &&
        currentDeps.isSubsetOf(dependenciesDiscoveredInEvaluation)
      ) ||
      currentDeps.size !== dependenciesDiscoveredInEvaluation.size
    ) {
      requiresReRun = true;
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
      requiresReRun = true;
    }

    this.storeManager.evaluatedNodes.set(nodeKey, {
      deps: dependenciesDiscoveredInEvaluation,
      frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
      discardedFrontierDependencies:
        discardedFrontierDependenciesDiscoveredInEvaluation,
      evaluationResult: evaluation,
    });

    return requiresReRun;
  }

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
        this.storeManager.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "error",
            err: FormulaError.ERROR,
            message: "Syntax error",
          },
        });
        break;
      }
      if (typeof content !== "string" || !content.startsWith("=")) {
        this.storeManager.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        });
        break;
      }

      // First, discover dependencies without full evaluation to detect cycles
      const discoveredDeps = this.discoverTransitiveDeps(nodeKey);
      discoveredDeps.add(nodeKey); // Include the starting node for cycle detection

      // Check for cycles using the discovered dependencies
      const cycleParticipants = this.findCycleParticipants(discoveredDeps);

      if (cycleParticipants.size > 0) {
        // cycle detected - mark all cycle participants and nodes that depend on them

        // Find all nodes that should be marked with cycle error
        // This includes cycle participants and any nodes that depend on them
        const nodesToMarkAsCycle = new Set(cycleParticipants);

        // Add any nodes that transitively depend on cycle participants
        for (const depNode of discoveredDeps) {
          if (!cycleParticipants.has(depNode)) {
            // Check if this node depends on any cycle participant
            const nodeDeps = this.discoverNodeDeps(depNode);
            for (const dep of nodeDeps) {
              if (cycleParticipants.has(dep)) {
                nodesToMarkAsCycle.add(depNode);
                break;
              }
            }
          }
        }

        // Also check if the original node should be marked
        if (!nodesToMarkAsCycle.has(nodeKey)) {
          // Check if the original node depends on any cycle participant
          const originalNodeDeps = this.discoverNodeDeps(nodeKey);
          for (const dep of originalNodeDeps) {
            if (cycleParticipants.has(dep)) {
              nodesToMarkAsCycle.add(nodeKey);
              break;
            }
          }
        }

        // Mark all nodes that should have cycle error
        for (const cycleNodeKey of nodesToMarkAsCycle) {
          this.storeManager.evaluatedNodes.set(cycleNodeKey, {
            deps:
              this.storeManager.evaluatedNodes.get(cycleNodeKey)?.deps ??
              new Set(),
            frontierDependencies:
              this.storeManager.evaluatedNodes.get(cycleNodeKey)
                ?.frontierDependencies ?? new Set(),
            discardedFrontierDependencies:
              this.storeManager.evaluatedNodes.get(cycleNodeKey)
                ?.discardedFrontierDependencies ?? new Set(),
            evaluationResult: {
              type: "error",
              err: FormulaError.CYCLE,
              message: "Cycle detected",
            },
          });
        }

        this.isEvaluating = false;
        return;
      }

      // No cycles detected, proceed with normal evaluation
      const allDeps = this.getTransitiveDeps(nodeKey);
      const sorted = this.topologicalSort(allDeps)?.reverse();

      if (!sorted) {
        // This shouldn't happen since we already checked for cycles above
        // But just in case, handle it gracefully
        this.storeManager.evaluatedNodes.set(nodeKey, {
          deps:
            this.storeManager.evaluatedNodes.get(nodeKey)?.deps ?? new Set(),
          frontierDependencies:
            this.storeManager.evaluatedNodes.get(nodeKey)
              ?.frontierDependencies ?? new Set(),
          discardedFrontierDependencies:
            this.storeManager.evaluatedNodes.get(nodeKey)
              ?.discardedFrontierDependencies ?? new Set(),
          evaluationResult: {
            type: "error",
            err: FormulaError.ERROR,
            message: "Unexpected topological sort failure",
          },
        });
        this.isEvaluating = false;
        return;
      }

      sorted.forEach((nodeKey) =>
        this.evaluateDependencyNode(nodeKey, cellAddress)
      );
      this.evaluateDependencyNode(nodeKey, cellAddress);

      const transitiveDeps2 = this.getTransitiveDeps(nodeKey);

      // the cells were potentially evaluated in the wrong order
      if (
        allDeps.size !== transitiveDeps2.size ||
        !allDeps.isSubsetOf(transitiveDeps2)
      ) {
        requiresReRun = true;
      }
    }
    this.isEvaluating = false;
  }

  convertScalarValueToCellValue(
    val: undefined | boolean | number | string
  ): CellValue {
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

    // maybe it is a spilled cell, we need to check the spilled values
    // the context is quite irrelevant, because the cells are "cached" if spilled is true
    const dummyContext: EvaluationContext = {
      currentCell: cellAddress,
      dependencies: new Set(),
      frontierDependencies: new Set(),
      discardedFrontierDependencies: new Set(),
    };
    const spilled = this.evaluateSpilled(cellAddress, dummyContext);
    if (spilled.isSpilled) {
      const result = spilled.result;
      if (!result) {
        return undefined;
      }
      if (result.type === "spilled-values") {
        return result.evaluate({ x: 0, y: 0 }, dummyContext);
      }
      return result;
    }

    const getEvaluatedNode = () => {
      return this.storeManager.evaluatedNodes.get(
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
      return result.evaluate({ x: 0, y: 0 }, dummyContext);
    }
    return result;
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.formulaEvaluator.isCellInTable(cellAddress);
  }
}
