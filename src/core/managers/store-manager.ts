import {
  type CellAddress,
  type EvaluatedDependencyNode,
  type EvaluationContext,
  type FunctionEvaluationResult,
  type SpilledValue,
  type SpreadsheetRange,
} from "../types";
import { isCellInRange } from "../utils";
import {
  dependencyNodeToKey,
  keyToDependencyNode,
} from "../utils/dependency-node-key";
import type { NamedExpressionManager } from "./named-expression-manager";

export class StoreManager {
  private evaluatedNodes: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey
     */
    string,
    EvaluatedDependencyNode
  > = new Map();

  /**
   * Index mapping frontier dependency nodeKey -> Set of nodeKeys that reference it in their frontierDependencies
   */
  private frontierDependencyIndex: Map<string, Set<string>> = new Map();

  /**
   * Index mapping discarded frontier dependency nodeKey -> Set of nodeKeys that reference it in their discardedFrontierDependencies
   */
  private discardedFrontierDependencyIndex: Map<string, Set<string>> =
    new Map();

  public spilledValues: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey for the origin cell
     */
    string,
    SpilledValue
  > = new Map();

  constructor(private namedExpressionManager: NamedExpressionManager) {}

  getSpillValue(cellAddress: CellAddress): SpilledValue | undefined {
    for (const spilledValue of this.spilledValues.values()) {
      if (spilledValue.origin.sheetName !== cellAddress.sheetName) {
        continue;
      }
      if (
        spilledValue.origin.colIndex === cellAddress.colIndex &&
        spilledValue.origin.rowIndex === cellAddress.rowIndex
      ) {
        return undefined;
      }
      if (isCellInRange(cellAddress, spilledValue.spillOnto)) {
        return spilledValue;
      }
    }
    return undefined;
  }

  getSpilledAddress(
    cellAddress: CellAddress,
    /**
     * if the spilled value is already available, we can use it to get the source address
     */
    passedSpilledValue?: SpilledValue
  ): { address: CellAddress; spillOffset: { x: number; y: number } } {
    const spilledValue = passedSpilledValue ?? this.getSpillValue(cellAddress);
    if (!spilledValue) {
      throw new Error("Cell is not spilled");
    }
    const offsetLeft = cellAddress.colIndex - spilledValue.origin.colIndex;
    const offsetTop = cellAddress.rowIndex - spilledValue.origin.rowIndex;
    const address: CellAddress = {
      ...cellAddress,
      colIndex: spilledValue.origin.colIndex + offsetLeft,
      rowIndex: spilledValue.origin.rowIndex + offsetTop,
    };
    if (offsetLeft === 0 && offsetTop === 0) {
      throw new Error(
        "Spilled value is the same as the cell address! The origin has a pre-calculated value that can be used"
      );
    }
    return { address, spillOffset: { x: offsetLeft, y: offsetTop } };
  }

  /**
   * During evaluation, we can't use the formula evaluator to evaluate a cell, because it will create a cycle.
   * This method can be used to "evaluate" a cell during evaluation, without creating a cycle.
   *
   * Internally this method will try to look up the evaluated result for the cell,
   * if it doesn't exist it will push the cell address to the dependency graph,
   * causing the engine to re-evaluate the dependency graph,
   * such that on the second evaluation the cell's evaluation result will be available.
   */
  evalTimeSafeEvaluateCell(
    cellAddress: CellAddress,
    context: EvaluationContext
  ): FunctionEvaluationResult | undefined {
    const spilled = this.getSpillValue(cellAddress);
    if (spilled) {
      const spillSource = this.getSpilledAddress(cellAddress, spilled);
      const spillOrigin = this.evalTimeSafeEvaluateCell(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return spillOrigin.evaluate(spillSource.spillOffset, context);
      }
    }
    const key = dependencyNodeToKey({
      address: cellAddress,
      sheetName: cellAddress.sheetName,
      workbookName: cellAddress.workbookName,
    });
    context.dependencies.add(key);
    const result = this.evaluatedNodes.get(key)?.evaluationResult;
    return result;
  }

  clearEvaluationCache(): void {
    this.evaluatedNodes.clear();
    this.spilledValues.clear();
    this.frontierDependencyIndex.clear();
    this.discardedFrontierDependencyIndex.clear();
  }

  /**
   * Helper method to add a node to the frontier dependency index
   */
  private addToFrontierIndex(frontierDep: string, nodeKey: string): void {
    let nodes = this.frontierDependencyIndex.get(frontierDep);
    if (!nodes) {
      nodes = new Set<string>();
      this.frontierDependencyIndex.set(frontierDep, nodes);
    }
    nodes.add(nodeKey);
  }

  /**
   * Helper method to remove a node from the frontier dependency index
   */
  private removeFromFrontierIndex(frontierDep: string, nodeKey: string): void {
    const nodes = this.frontierDependencyIndex.get(frontierDep);
    if (nodes) {
      nodes.delete(nodeKey);
      if (nodes.size === 0) {
        this.frontierDependencyIndex.delete(frontierDep);
      }
    }
  }

  /**
   * Helper method to add a node to the discarded frontier dependency index
   */
  private addToDiscardedIndex(discardedDep: string, nodeKey: string): void {
    let nodes = this.discardedFrontierDependencyIndex.get(discardedDep);
    if (!nodes) {
      nodes = new Set<string>();
      this.discardedFrontierDependencyIndex.set(discardedDep, nodes);
    }
    nodes.add(nodeKey);
  }

  /**
   * Helper method to remove a node from the discarded frontier dependency index
   */
  private removeFromDiscardedIndex(
    discardedDep: string,
    nodeKey: string
  ): void {
    const nodes = this.discardedFrontierDependencyIndex.get(discardedDep);
    if (nodes) {
      nodes.delete(nodeKey);
      if (nodes.size === 0) {
        this.discardedFrontierDependencyIndex.delete(discardedDep);
      }
    }
  }

  /**
   * The nodeKey is producing spilled values over the range so the candidate should be restored
   * as a dependency for dependencies intersecting with the range
   */
  restoreFrontierCandidate(nodeKey: string, range: SpreadsheetRange): void {
    const nodesWithDiscardedDep =
      this.discardedFrontierDependencyIndex.get(nodeKey);
    if (!nodesWithDiscardedDep) {
      return;
    }

    // Clone the set to avoid iterator invalidation during removal
    const nodesToCheck = [...nodesWithDiscardedDep];

    for (const key of nodesToCheck) {
      const evalNode = this.evaluatedNodes.get(key);
      if (!evalNode || !evalNode.discardedFrontierDependencies?.has(nodeKey)) {
        continue;
      }

      const depNode = keyToDependencyNode(key);
      if (isCellInRange(depNode.address, range)) {
        evalNode.discardedFrontierDependencies.delete(nodeKey);
        this.removeFromDiscardedIndex(nodeKey, key);
      }
    }
  }

  /**
   * The nodeKey is not producing spilled values so the candidate should be discarded
   * as a dependency
   */
  discardFrontierCandidate(nodeKey: string) {
    const nodesWithFrontierDep = this.frontierDependencyIndex.get(nodeKey);
    if (!nodesWithFrontierDep) {
      return;
    }

    // Clone the set to avoid iterator invalidation during modification
    const nodesToUpdate = [...nodesWithFrontierDep];

    for (const key of nodesToUpdate) {
      const evalNode = this.evaluatedNodes.get(key);
      if (!evalNode || !evalNode.frontierDependencies?.has(nodeKey)) {
        continue;
      }

      if (!evalNode.discardedFrontierDependencies) {
        evalNode.discardedFrontierDependencies = new Set<string>();
      }
      evalNode.discardedFrontierDependencies.add(nodeKey);
      this.addToDiscardedIndex(nodeKey, key);
    }
  }

  setEvaluatedNode(nodeKey: string, node: EvaluatedDependencyNode): void {
    const currentNode = this.evaluatedNodes.get(nodeKey);

    // Remove old frontier dependencies from index
    if (currentNode?.frontierDependencies) {
      for (const dep of currentNode.frontierDependencies) {
        this.removeFromFrontierIndex(dep, nodeKey);
      }
    }

    // Remove old discarded frontier dependencies from index
    if (currentNode?.discardedFrontierDependencies) {
      for (const dep of currentNode.discardedFrontierDependencies) {
        this.removeFromDiscardedIndex(dep, nodeKey);
      }
    }

    // Add new frontier dependencies to index
    if (node.frontierDependencies) {
      for (const dep of node.frontierDependencies) {
        this.addToFrontierIndex(dep, nodeKey);
      }
    }

    // Add new discarded frontier dependencies to index
    if (node.discardedFrontierDependencies) {
      for (const dep of node.discardedFrontierDependencies) {
        this.addToDiscardedIndex(dep, nodeKey);
      }
    }

    this.evaluatedNodes.set(nodeKey, node);
  }

  getEvaluatedNode(nodeKey: string): EvaluatedDependencyNode | undefined {
    return this.evaluatedNodes.get(nodeKey);
  }

  getEvaluatedNodes(): Map<string, EvaluatedDependencyNode> {
    return this.evaluatedNodes;
  }
}
