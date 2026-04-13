import {
  getAstNodeSnapshotId,
  type CacheManagerSnapshot,
  type DependencyManagerSnapshot,
  type NodeSnapshotId,
  type SerializedAstNodeSnapshot,
  type SerializedResourceNodeSnapshot,
  type SerializedDependencyNodeSnapshot,
  type SerializedEmptyCellNodeSnapshot,
  type SerializedRangeNodeSnapshot,
  type SerializedSpillMetaNodeSnapshot,
  type SerializedCellValueNodeSnapshot,
} from "../engine-snapshot";
import {
  eligibleKeysForContext,
  getContextDependencyKey,
  type ContextDependency,
} from "../../evaluator/evaluation-context";
import {
  type CellAddress,
  type EvaluationOrder,
  type RangeAddress,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpilledValue,
} from "../types";
import {
  cellAddressToKey,
  checkRangeIntersection,
  isCellInRange,
  keyToCellAddress,
} from "../utils";

import { AstEvaluationNode } from "../../evaluator/dependency-nodes/ast-evaluation-node";
import { CellValueNode } from "../../evaluator/dependency-nodes/cell-value-node";
import { EmptyCellEvaluationNode } from "../../evaluator/dependency-nodes/empty-cell-evaluation-node";
import { SpillMetaNode } from "../../evaluator/dependency-nodes/spill-meta-node";
import { RangeEvaluationNode } from "../../evaluator/range-evaluation-node";
import type { ASTNode } from "../../parser/ast";
import {
  astToString,
  normalizeSerializedCellValue,
} from "../../parser/formatter";
import { CacheManager } from "./cache-manager";
import type {
  CellNodeKeyDictionary,
  CellNodeType,
  DependencyNode,
} from "./dependency-node";
import { WorkbookManager } from "./workbook-manager";
import { VirtualCellValueNode } from "../../evaluator/dependency-nodes/virtual-cell-value-node";
import { parseFormula } from "../../parser/parser";
import type { EvaluationManager } from "./evaluation-manager";
import type { MutationInvalidation, RemovedScope } from "../commands/types";
import { ResourceDependencyNode } from "../../evaluator/dependency-nodes/resource-dependency-node";

export interface DependencyTreeNode {
  type: "cell" | "range" | "empty";
  circular?: boolean;
  key: string;
  directDepsUpdated?: boolean;
  resolved?: boolean;
  canResolve: boolean;
  resultType:
    | "awaiting-evaluation"
    | "spilled-values"
    | "value"
    | "range"
    | "error"
    | "does-not-spill"
    | "resource";
  deps?: DependencyTreeNode[];
  frontierDependencies?: DependencyTreeNode[];
  self?: boolean;
  _debug?: {
    rawFrontierDependencies?: string[];
    discardedFrontierDependencies?: string[];
    activeFrontierDependencies?: string[];
  };
}

type FrontierWatcherNode = RangeEvaluationNode | EmptyCellEvaluationNode;

/**
 * The DependencyManager is responsible for storing the evaluated values and their dependencies.
 */
export class DependencyManager {
  /**
   * The dependency graph AKA cellNodes
   */
  private cellNodes: Map<
    /**
     * key is the cell key, from cellAddressToKey
     */
    string,
    CellValueNode
  > = new Map();

  private spillMetaNodes: Map<string, SpillMetaNode> = new Map();

  private emptyCells: Map<string, EmptyCellEvaluationNode> = new Map();

  private virtualCellValueNodes: Map<string, VirtualCellValueNode> = new Map();

  /**
   * registry of spilled values
   */
  private _spilledValues: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey for the origin cell
     */
    string,
    SpilledValue
  > = new Map();

  /**
   * Key is workbook:sheetName:rangeKey, e.g. Workbook1:Sheet1:A1:D10, from rangeAddressToKey
   */
  private ranges: Map<string, RangeEvaluationNode> = new Map();

  private resourceNodes: Map<string, ResourceDependencyNode> = new Map();

  private hardDependents: Map<DependencyNode, Set<DependencyNode>> = new Map();

  private frontierDependents: Map<DependencyNode, Set<DependencyNode>> =
    new Map();

  private registeredDependencies: Map<
    DependencyNode,
    {
      hard: Set<DependencyNode>;
      frontier: Set<DependencyNode>;
    }
  > = new Map();

  private coverageWatchersBySheet: Map<string, Set<FrontierWatcherNode>> =
    new Map();

  private frontierWatchersBySheet: Map<string, Set<FrontierWatcherNode>> =
    new Map();

  constructor(
    private cacheManager: CacheManager,
    private workbookManager: WorkbookManager
  ) {}

  private getSheetWatcherKey(address: {
    workbookName: string;
    sheetName: string;
  }): string {
    return `${address.workbookName}:${address.sheetName}`;
  }

  private addReverseEdge(
    map: Map<DependencyNode, Set<DependencyNode>>,
    dependency: DependencyNode,
    dependent: DependencyNode
  ) {
    const dependents = map.get(dependency) ?? new Set<DependencyNode>();
    dependents.add(dependent);
    map.set(dependency, dependents);
  }

  private removeReverseEdge(
    map: Map<DependencyNode, Set<DependencyNode>>,
    dependency: DependencyNode,
    dependent: DependencyNode
  ) {
    const dependents = map.get(dependency);
    if (!dependents) {
      return;
    }
    dependents.delete(dependent);
    if (dependents.size === 0) {
      map.delete(dependency);
    }
  }

  private getWatcherRange(node: FrontierWatcherNode): RangeAddress {
    return node instanceof RangeEvaluationNode
      ? node.address
      : node.getFrontierRange();
  }

  private addWatcher(
    map: Map<string, Set<FrontierWatcherNode>>,
    node: FrontierWatcherNode
  ) {
    const watcherKey = this.getSheetWatcherKey(this.getWatcherRange(node));
    const watchers = map.get(watcherKey) ?? new Set<FrontierWatcherNode>();
    watchers.add(node);
    map.set(watcherKey, watchers);
  }

  private removeWatcher(
    map: Map<string, Set<FrontierWatcherNode>>,
    node: FrontierWatcherNode
  ) {
    const watcherKey = this.getSheetWatcherKey(this.getWatcherRange(node));
    const watchers = map.get(watcherKey);
    if (!watchers) {
      return;
    }
    watchers.delete(node);
    if (watchers.size === 0) {
      map.delete(watcherKey);
    }
  }

  private registerWatcherNode(node: FrontierWatcherNode) {
    this.addWatcher(this.coverageWatchersBySheet, node);
    this.addWatcher(this.frontierWatchersBySheet, node);
  }

  private unregisterWatcherNode(node: FrontierWatcherNode) {
    this.removeWatcher(this.coverageWatchersBySheet, node);
    this.removeWatcher(this.frontierWatchersBySheet, node);
  }

  private isWatcherNodeResolved(node: FrontierWatcherNode): boolean {
    return node instanceof RangeEvaluationNode
      ? node.result.type !== "awaiting-evaluation"
      : node.evaluationResult.type !== "awaiting-evaluation";
  }

  private getAllPersistentNodes(): DependencyNode[] {
    const astNodes: AstEvaluationNode[] = [];
    for (const astEntries of this.asts.values()) {
      for (const astEntry of astEntries.entries.values()) {
        astNodes.push(astEntry.evalNode);
      }
    }

    return [
      ...this.cellNodes.values(),
      ...this.spillMetaNodes.values(),
      ...this.emptyCells.values(),
      ...this.ranges.values(),
      ...this.resourceNodes.values(),
      ...astNodes,
    ];
  }

  private getExistingCellValueNode(nodeKey: string): CellValueNode | undefined {
    return this.cellNodes.get(nodeKey);
  }

  private getExistingSpillMetaNode(nodeKey: string): SpillMetaNode | undefined {
    return this.spillMetaNodes.get(nodeKey);
  }

  private getExistingEmptyCellNode(
    nodeKey: string
  ): EmptyCellEvaluationNode | undefined {
    return this.emptyCells.get(nodeKey);
  }

  private collectExistingNodesForCell(address: CellAddress): DependencyNode[] {
    const baseKey = cellAddressToKey(address);
    const nodes: Array<
      CellValueNode | EmptyCellEvaluationNode | SpillMetaNode | undefined
    > = [
      this.getExistingCellValueNode(baseKey),
      this.getExistingEmptyCellNode(baseKey.replace(/^cell-value:/, "empty:")),
      this.getExistingSpillMetaNode(
        baseKey.replace(/^cell-value:/, "spill-meta:")
      ),
    ];
    return nodes.filter(
      (
        node
      ): node is CellValueNode | EmptyCellEvaluationNode | SpillMetaNode =>
        node !== undefined
    );
  }

  private collectExistingAstNodesForCell(
    address: CellAddress
  ): AstEvaluationNode[] {
    const cellNode = this.getExistingCellValueNode(cellAddressToKey(address));
    if (!cellNode) {
      return [];
    }

    const collected = new Set<AstEvaluationNode>();
    const stack = Array.from(cellNode.getDependencies()).filter(
      (dependency): dependency is AstEvaluationNode =>
        dependency instanceof AstEvaluationNode
    );

    while (stack.length > 0) {
      const dependency = stack.pop();
      if (!dependency) {
        continue;
      }

      if (collected.has(dependency)) {
        continue;
      }
      collected.add(dependency);

      for (const nestedDependency of dependency.getDependencies()) {
        if (nestedDependency instanceof AstEvaluationNode) {
          stack.push(nestedDependency);
        }
      }
    }

    return Array.from(collected);
  }

  private collectExistingAstNodesForCells(
    addresses: CellAddress[]
  ): AstEvaluationNode[] {
    const collected = new Set<AstEvaluationNode>();

    for (const address of addresses) {
      for (const astNode of this.collectExistingAstNodesForCell(address)) {
        collected.add(astNode);
      }
    }

    return Array.from(collected);
  }

  private astNodeHasExternalDependents(
    node: AstEvaluationNode,
    formulaAstNodes: Set<AstEvaluationNode>,
    removedDependents: Set<DependencyNode>
  ): boolean {
    for (const dependent of this.getNodeDependents(node)) {
      if (removedDependents.has(dependent)) {
        continue;
      }

      if (dependent instanceof AstEvaluationNode && formulaAstNodes.has(dependent)) {
        continue;
      }

      return true;
    }

    return false;
  }

  private collectOrphanedOldFormulaAstNodesForCell(
    address: CellAddress
  ): AstEvaluationNode[] {
    return this.collectOrphanedOldFormulaAstNodesForCells([address]);
  }

  private collectOrphanedOldFormulaAstNodesForCells(
    addresses: CellAddress[]
  ): AstEvaluationNode[] {
    const astNodes = this.collectExistingAstNodesForCells(addresses);
    if (astNodes.length === 0) {
      return [];
    }

    const formulaAstNodes = new Set(astNodes);
    const removedDependents = new Set<DependencyNode>(
      addresses.flatMap((address) => this.collectExistingNodesForCell(address))
    );
    const keptAstNodes = new Set<AstEvaluationNode>();
    const stack = astNodes.filter((node) =>
      this.astNodeHasExternalDependents(node, formulaAstNodes, removedDependents)
    );

    while (stack.length > 0) {
      const node = stack.pop();
      if (!node || keptAstNodes.has(node)) {
        continue;
      }
      keptAstNodes.add(node);

      for (const dependency of node.getDependencies()) {
        if (dependency instanceof AstEvaluationNode && formulaAstNodes.has(dependency)) {
          stack.push(dependency);
        }
      }
    }

    return astNodes.filter((node) => !keptAstNodes.has(node));
  }

  private removeAstNodeFromCache(node: AstEvaluationNode): void {
    const astEntries = this.asts.get(node.key);
    if (!astEntries) {
      return;
    }

    for (const [contextKey, astEntry] of Array.from(astEntries.entries.entries())) {
      if (astEntry.evalNode === node) {
        astEntries.entries.delete(contextKey);
      }
    }

    if (astEntries.entries.size === 0) {
      this.asts.delete(node.key);
    }
  }

  private collectSpillOriginsAffectingCell(
    address: CellAddress
  ): Set<DependencyNode> {
    const affected = new Set<DependencyNode>();

    for (const [spillOriginKey, spilledValue] of this._spilledValues.entries()) {
      if (
        spilledValue.origin.workbookName !== address.workbookName ||
        spilledValue.origin.sheetName !== address.sheetName
      ) {
        continue;
      }

      const isOriginCell =
        spilledValue.origin.colIndex === address.colIndex &&
        spilledValue.origin.rowIndex === address.rowIndex;
      if (isOriginCell || !isCellInRange(address, spilledValue.spillOnto)) {
        continue;
      }

      const cellValueNode = this.cellNodes.get(`cell-value:${spillOriginKey}`);
      if (cellValueNode) {
        affected.add(cellValueNode);
      }

      const spillMetaNode = this.spillMetaNodes.get(
        `spill-meta:${spillOriginKey}`
      );
      if (spillMetaNode) {
        affected.add(spillMetaNode);
      }
    }

    return affected;
  }

  private getLinkedSpillMetaNode(
    node: CellValueNode | SpillMetaNode
  ): SpillMetaNode | undefined {
    if (node instanceof SpillMetaNode) {
      return node;
    }

    return this.spillMetaNodes.get(
      node.key.replace(/^cell-value:/, "spill-meta:")
    );
  }

  private getLinkedCellValueNode(
    node: CellValueNode | SpillMetaNode
  ): CellValueNode | undefined {
    if (node instanceof CellValueNode) {
      return node;
    }

    return this.cellNodes.get(node.key.replace(/^spill-meta:/, "cell-value:"));
  }

  public registerNode(node: DependencyNode): void {
    this.unregisterNode(node);

    const hardDependencies = new Set(node.getDependencies());
    const frontierDependencies =
      node instanceof RangeEvaluationNode || node instanceof EmptyCellEvaluationNode
        ? new Set(node.getFrontierDependencies())
        : new Set<DependencyNode>();

    this.registeredDependencies.set(node, {
      hard: hardDependencies,
      frontier: frontierDependencies,
    });

    for (const dependency of hardDependencies) {
      this.addReverseEdge(this.hardDependents, dependency, node);
    }

    for (const dependency of frontierDependencies) {
      this.addReverseEdge(this.frontierDependents, dependency, node);
    }

    if (
      (node instanceof RangeEvaluationNode ||
        node instanceof EmptyCellEvaluationNode) &&
      this.isWatcherNodeResolved(node)
    ) {
      this.registerWatcherNode(node);
    }
  }

  public unregisterNode(node: DependencyNode): void {
    const registration = this.registeredDependencies.get(node);
    if (registration) {
      for (const dependency of registration.hard) {
        this.removeReverseEdge(this.hardDependents, dependency, node);
      }

      for (const dependency of registration.frontier) {
        this.removeReverseEdge(this.frontierDependents, dependency, node);
      }

      this.registeredDependencies.delete(node);
    }

    if (node instanceof RangeEvaluationNode || node instanceof EmptyCellEvaluationNode) {
      this.unregisterWatcherNode(node);
    }
  }

  private rebuildRuntimeIndexes() {
    this.hardDependents.clear();
    this.frontierDependents.clear();
    this.registeredDependencies.clear();
    this.coverageWatchersBySheet.clear();
    this.frontierWatchersBySheet.clear();

    for (const node of this.getAllPersistentNodes()) {
      this.registerNode(node);
    }
  }

  public get spilledValues(): IterableIterator<SpilledValue> {
    return this._spilledValues.values();
  }

  isSpillOrigin(cellAddress: CellAddress): boolean {
    for (const spilledValue of this._spilledValues.values()) {
      if (
        spilledValue.origin.sheetName !== cellAddress.sheetName ||
        spilledValue.origin.workbookName !== cellAddress.workbookName
      ) {
        continue;
      }
      if (
        spilledValue.origin.colIndex === cellAddress.colIndex &&
        spilledValue.origin.rowIndex === cellAddress.rowIndex
      ) {
        return true;
      }
    }
    return false;
  }

  getSpillValue(cellAddress: CellAddress): SpilledValue | undefined {
    for (const spilledValue of this._spilledValues.values()) {
      if (
        spilledValue.origin.sheetName !== cellAddress.sheetName ||
        spilledValue.origin.workbookName !== cellAddress.workbookName
      ) {
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

  clearEvaluationCache(): void {
    this.cacheManager.clear();
    this.cellNodes.clear();
    this.emptyCells.clear();
    this.virtualCellValueNodes.clear();
    this.asts.clear();
    this.spillMetaNodes.clear();
    this.ranges.clear();
    this.resourceNodes.clear();
    this._spilledValues.clear();
    this.hardDependents.clear();
    this.frontierDependents.clear();
    this.registeredDependencies.clear();
    this.coverageWatchersBySheet.clear();
    this.frontierWatchersBySheet.clear();
  }

  toSnapshot(evaluationManager: EvaluationManager): {
    dependency: DependencyManagerSnapshot;
    cache: CacheManagerSnapshot;
  } {
    const isNodeSnapshotEligible = (
      node:
        | CellValueNode
        | SpillMetaNode
        | EmptyCellEvaluationNode
        | RangeEvaluationNode
        | AstEvaluationNode
        | ResourceDependencyNode
    ) => {
      if (node instanceof ResourceDependencyNode) {
        return true;
      }
      if (node instanceof RangeEvaluationNode) {
        return node.result.type !== "awaiting-evaluation";
      }
      return node.evaluationResult.type !== "awaiting-evaluation";
    };

    const astNodes = new Set<AstEvaluationNode>();
    for (const astEntries of this.asts.values()) {
      for (const { evalNode } of astEntries.entries.values()) {
        if (isNodeSnapshotEligible(evalNode)) {
          astNodes.add(evalNode);
        }
      }
    }

    const snapshotSeedNodes: Array<
      | CellValueNode
      | SpillMetaNode
      | EmptyCellEvaluationNode
      | RangeEvaluationNode
      | AstEvaluationNode
      | ResourceDependencyNode
    > = [
      ...Array.from(this.cellNodes.values()).filter(isNodeSnapshotEligible),
      ...Array.from(this.spillMetaNodes.values()).filter(isNodeSnapshotEligible),
      ...Array.from(this.emptyCells.values()).filter(isNodeSnapshotEligible),
      ...Array.from(this.ranges.values()).filter(isNodeSnapshotEligible),
      ...Array.from(this.resourceNodes.values()).filter(isNodeSnapshotEligible),
      ...Array.from(astNodes.values()),
    ];

    const resolvedNodesSet = new Set<
      | CellValueNode
      | SpillMetaNode
      | EmptyCellEvaluationNode
      | RangeEvaluationNode
      | AstEvaluationNode
      | ResourceDependencyNode
    >();
    const stack = [...snapshotSeedNodes];

    while (stack.length > 0) {
      const node = stack.pop();
      if (!node || resolvedNodesSet.has(node)) {
        continue;
      }

      resolvedNodesSet.add(node);

      if (
        node instanceof CellValueNode &&
        node.spillMeta &&
        isNodeSnapshotEligible(node.spillMeta)
      ) {
        stack.push(node.spillMeta);
      }

      for (const dependency of node.getDependencies()) {
        if (!isNodeSnapshotEligible(dependency)) {
          continue;
        }
        stack.push(dependency);
      }
    }

    const resolvedNodes = Array.from(resolvedNodesSet);

    const allNodeSnapshotIds = new Map<DependencyNode, NodeSnapshotId>();
    for (const node of resolvedNodes) {
      if (node instanceof AstEvaluationNode) {
        allNodeSnapshotIds.set(node, getAstNodeSnapshotId(node));
      } else {
        allNodeSnapshotIds.set(node, node.key);
      }
    }

    const getAllNodeSnapshotId = (node: DependencyNode): NodeSnapshotId => {
      const snapshotId = allNodeSnapshotIds.get(node);
      if (!snapshotId) {
        throw new Error(`Missing snapshot id for dependency node ${node.key}`);
      }
      return snapshotId;
    };

    const nodeSnapshots: SerializedDependencyNodeSnapshot[] = [];
    const includedSnapshotIds = new Set<NodeSnapshotId>();

    for (const node of resolvedNodes) {
      const snapshot = this.serializeNodeSnapshot(
        node,
        evaluationManager,
        getAllNodeSnapshotId,
        allNodeSnapshotIds
      );
      if (!snapshot) {
        continue;
      }
      includedSnapshotIds.add(snapshot.snapshotId);
      nodeSnapshots.push(snapshot);
    }

    const filteredNodeSnapshots = nodeSnapshots.map((snapshot) => {
      const dependencies = snapshot.dependencies.filter((dependency) =>
        includedSnapshotIds.has(dependency)
      );

      if (snapshot.kind === "cell-value") {
        return {
          ...snapshot,
          dependencies,
          spillMetaSnapshotId:
            snapshot.spillMetaSnapshotId &&
            includedSnapshotIds.has(snapshot.spillMetaSnapshotId)
              ? snapshot.spillMetaSnapshotId
              : undefined,
        };
      }

      return {
        ...snapshot,
        dependencies,
      };
    });

    const dependency: DependencyManagerSnapshot = {
      nodes: filteredNodeSnapshots,
      spilledValues: Array.from(this._spilledValues.entries()),
    };

    const cache = this.cacheManager.toSnapshot((node) => {
      const snapshotId = allNodeSnapshotIds.get(node);
      if (!snapshotId || !includedSnapshotIds.has(snapshotId)) {
        return undefined;
      }
      return snapshotId;
    });

    return { dependency, cache };
  }

  restoreFromSnapshot(
    snapshots: {
      dependency: DependencyManagerSnapshot;
      cache: CacheManagerSnapshot;
    },
    evaluationManager: EvaluationManager
  ) {
    this.clearEvaluationCache();

    const nodesBySnapshotId = new Map<NodeSnapshotId, DependencyNode>();

    for (const snapshot of snapshots.dependency.nodes) {
      const node = this.createNodeFromSnapshot(snapshot);
      nodesBySnapshotId.set(snapshot.snapshotId, node);
    }

    const resolveNodeSnapshotId = (snapshotId: NodeSnapshotId) =>
      nodesBySnapshotId.get(snapshotId);
    const resolveRequiredNodeSnapshotId = (
      snapshotId: NodeSnapshotId
    ): DependencyNode => {
      const node = resolveNodeSnapshotId(snapshotId);
      if (!node) {
        throw new Error(`Unknown node snapshot id: ${snapshotId}`);
      }
      return node;
    };

    for (const snapshot of snapshots.dependency.nodes) {
      const dependencies = new Set(
        snapshot.dependencies
          .map(resolveNodeSnapshotId)
          .filter((node): node is DependencyNode => node !== undefined)
      );

      if (snapshot.kind === "cell-value") {
        const node = nodesBySnapshotId.get(snapshot.snapshotId) as CellValueNode;
        node.restoreResolvedSnapshot({
          dependencies,
          evaluationResult:
            evaluationManager.deserializeSingleEvaluationResultSnapshot(
              snapshot.evaluationResult,
              resolveRequiredNodeSnapshotId
            ),
        });
        if (snapshot.spillMetaSnapshotId) {
          const spillMetaNode = resolveNodeSnapshotId(
            snapshot.spillMetaSnapshotId
          );
          if (spillMetaNode instanceof SpillMetaNode) {
            node.setSpillMetaNode(spillMetaNode);
          }
        }
        continue;
      }

      if (snapshot.kind === "spill-meta") {
        const node = nodesBySnapshotId.get(snapshot.snapshotId) as SpillMetaNode;
        node.restoreResolvedSnapshot({
          dependencies,
          evaluationResult:
            evaluationManager.deserializeSpillMetaEvaluationResultSnapshot(
              snapshot.evaluationResult,
              resolveRequiredNodeSnapshotId
            ),
        });
        continue;
      }

      if (snapshot.kind === "empty") {
        const node = nodesBySnapshotId.get(
          snapshot.snapshotId
        ) as EmptyCellEvaluationNode;
        node.restoreResolvedSnapshot({
          dependencies,
          evaluationResult:
            evaluationManager.deserializeSingleEvaluationResultSnapshot(
              snapshot.evaluationResult,
              resolveRequiredNodeSnapshotId
            ),
        });
        continue;
      }

      if (snapshot.kind === "range") {
        const node = nodesBySnapshotId.get(
          snapshot.snapshotId
        ) as RangeEvaluationNode;
        node.restoreResolvedSnapshot({
          dependencies,
          result: evaluationManager.deserializeEvaluateAllCellsResultSnapshot(
            snapshot.result,
            resolveRequiredNodeSnapshotId
          ),
        });
        continue;
      }

      if (snapshot.kind === "resource") {
        continue;
      }

      const node = nodesBySnapshotId.get(snapshot.snapshotId) as AstEvaluationNode;
      node.restoreResolvedSnapshot({
        dependencies,
        evaluationResult:
          evaluationManager.deserializeFunctionEvaluationResultSnapshot(
            snapshot.evaluationResult,
            resolveRequiredNodeSnapshotId
          ),
      });
      this.saveAstNode(node, snapshot.contextDependency);
    }

    this._spilledValues = new Map(snapshots.dependency.spilledValues);
    this.rebuildRuntimeIndexes();
    this.cacheManager.restoreFromSnapshot(
      snapshots.cache,
      resolveNodeSnapshotId
    );
  }

  private serializeNodeSnapshot(
    node:
      | CellValueNode
      | SpillMetaNode
      | EmptyCellEvaluationNode
      | RangeEvaluationNode
      | AstEvaluationNode
      | ResourceDependencyNode,
    evaluationManager: EvaluationManager,
    getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId,
    allNodeSnapshotIds: Map<DependencyNode, NodeSnapshotId>
  ): SerializedDependencyNodeSnapshot | undefined {
    const snapshotId = getNodeSnapshotId(node);
    const dependencies: NodeSnapshotId[] = [];
    for (const dependency of node.getDependencies()) {
      const dependencySnapshotId = allNodeSnapshotIds.get(dependency);
      if (!dependencySnapshotId) {
        return undefined;
      }
      dependencies.push(dependencySnapshotId);
    }

    if (node instanceof CellValueNode) {
      const spillMetaSnapshotId = node.spillMeta
        ? allNodeSnapshotIds.get(node.spillMeta)
        : undefined;
      if (node.spillMeta && !spillMetaSnapshotId) {
        return undefined;
      }

      const evaluationResult =
        evaluationManager.serializeSingleEvaluationResultSnapshot(
          node.evaluationResult,
          getNodeSnapshotId
        );
      if (!evaluationResult) {
        return undefined;
      }
      return {
        kind: "cell-value",
        snapshotId,
        key: node.key,
        dependencies,
        evaluationResult,
        spillMetaSnapshotId,
      };
    }

    if (node instanceof SpillMetaNode) {
      const evaluationResult =
        evaluationManager.serializeSpillMetaEvaluationResultSnapshot(
          node.evaluationResult,
          {
            sourceNode: node,
            getNodeSnapshotId,
          }
        );
      if (!evaluationResult) {
        return undefined;
      }
      return {
        kind: "spill-meta",
        snapshotId,
        key: node.key,
        dependencies,
        evaluationResult,
      };
    }

    if (node instanceof EmptyCellEvaluationNode) {
      const evaluationResult =
        evaluationManager.serializeSingleEvaluationResultSnapshot(
          node.evaluationResult,
          getNodeSnapshotId
        );
      if (!evaluationResult) {
        return undefined;
      }
      return {
        kind: "empty",
        snapshotId,
        key: node.key,
        dependencies,
        evaluationResult,
      };
    }

    if (node instanceof RangeEvaluationNode) {
      const result = evaluationManager.serializeEvaluateAllCellsResultSnapshot(
        node.result,
        getNodeSnapshotId
      );
      if (!result) {
        return undefined;
      }
      return {
        kind: "range",
        snapshotId,
        key: node.key,
        dependencies,
        result,
      };
    }

    if (node instanceof ResourceDependencyNode) {
      return {
        kind: "resource",
        snapshotId,
        key: node.key,
        dependencies,
      };
    }

    const evaluationResult =
      evaluationManager.serializeFunctionEvaluationResultSnapshot(
        node.evaluationResult,
        {
          sourceNode: node,
          getNodeSnapshotId,
        }
      );
    if (!evaluationResult) {
      return undefined;
    }
    return {
      kind: "ast",
      snapshotId,
      key: node.key,
      dependencies,
      contextDependency: node.getContextDependency(),
      evaluationResult,
    };
  }

  private createNodeFromSnapshot(
    snapshot: SerializedDependencyNodeSnapshot
  ): DependencyNode {
    if (snapshot.kind === "cell-value") {
      const node = new CellValueNode(snapshot.key);
      this.cellNodes.set(snapshot.key, node);
      return node;
    }

    if (snapshot.kind === "spill-meta") {
      const node = new SpillMetaNode(snapshot.key);
      this.spillMetaNodes.set(snapshot.key, node);
      return node;
    }

    if (snapshot.kind === "empty") {
      const node = new EmptyCellEvaluationNode(
        snapshot.key,
        this,
        this.workbookManager,
        { skipInitialBuild: true }
      );
      this.emptyCells.set(snapshot.key, node);
      return node;
    }

    if (snapshot.kind === "range") {
      const node = new RangeEvaluationNode(
        snapshot.key,
        this,
        this.workbookManager,
        { skipInitialBuild: true }
      );
      this.ranges.set(snapshot.key, node);
      return node;
    }

    if (snapshot.kind === "resource") {
      const node = new ResourceDependencyNode(snapshot.key);
      this.resourceNodes.set(snapshot.key, node);
      return node;
    }

    const node = new AstEvaluationNode(
      parseFormula(snapshot.key.slice(4)),
      snapshot.contextDependency
    );
    return node;
  }

  setSpilledValue(nodeKey: string, spilledValue: SpilledValue): void {
    this._spilledValues.set(nodeKey.replace(/^[^:]+:/, ""), spilledValue);
  }

  getSpilledValue(nodeKey: string): SpilledValue | undefined {
    return this._spilledValues.get(nodeKey.replace(/^[^:]+:/, ""));
  }

  deleteSpilledValue(nodeKey: string): void {
    this._spilledValues.delete(nodeKey.replace(/^[^:]+:/, ""));
  }

  getEmptyCellNode(nodeKey: string): EmptyCellEvaluationNode {
    if (!nodeKey.startsWith("empty:")) {
      throw new Error("Invalid empty cell node key: " + nodeKey);
    }
    if (!this.emptyCells.has(nodeKey)) {
      const node = new EmptyCellEvaluationNode(
        nodeKey,
        this,
        this.workbookManager
      );
      this.emptyCells.set(nodeKey, node);
      return node;
    }
    return this.emptyCells.get(nodeKey)!;
  }

  getSpillMetaNode(nodeKey: string): SpillMetaNode {
    if (!nodeKey.startsWith("spill-meta:")) {
      throw new Error("Invalid spill meta node key: " + nodeKey);
    }
    if (!this.spillMetaNodes.has(nodeKey)) {
      const node = new SpillMetaNode(nodeKey);
      this.spillMetaNodes.set(nodeKey, node);
      return node;
    }
    return this.spillMetaNodes.get(nodeKey)!;
  }

  getCellValueNode(nodeKey: string): CellValueNode {
    if (!nodeKey.startsWith("cell-value:")) {
      throw new Error("Invalid cell value node key: " + nodeKey);
    }
    if (!this.cellNodes.has(nodeKey)) {
      const node = new CellValueNode(nodeKey);
      this.cellNodes.set(nodeKey, node);
      return node;
    }
    return this.cellNodes.get(nodeKey)!;
  }

  getResourceNode(resourceKey: string): ResourceDependencyNode {
    if (!this.resourceNodes.has(resourceKey)) {
      const node = new ResourceDependencyNode(resourceKey);
      this.resourceNodes.set(resourceKey, node);
      return node;
    }
    return this.resourceNodes.get(resourceKey)!;
  }

  getCellValueOrEmptyCellNode(
    nodeKey: string
  ): CellValueNode | EmptyCellEvaluationNode {
    const cellAddress = keyToCellAddress(nodeKey);

    const emptyKey = nodeKey.replace(/^[^:]+:/, "empty:");
    const cellValueKey = nodeKey.replace(/^[^:]+:/, "cell-value:");

    if (this.workbookManager.isCellEmpty(cellAddress)) {
      return this.getEmptyCellNode(emptyKey);
    }

    return this.getCellValueNode(cellValueKey);
  }

  getSpillMetaOrEmptySpillMetaNode(
    nodeKey: string
  ): SpillMetaNode | EmptyCellEvaluationNode {
    const cellAddress = keyToCellAddress(nodeKey);

    const emptyKey = nodeKey.replace(/^[^:]+:/, "empty:");
    const spillMetaKey = nodeKey.replace(/^[^:]+:/, "spill-meta:");

    if (this.workbookManager.isCellEmpty(cellAddress)) {
      return this.getEmptyCellNode(emptyKey);
    }

    return this.getSpillMetaNode(spillMetaKey);
  }

  getVirtualCellValueNode(
    cellAddress: CellAddress,
    cellValue: SerializedCellValue
  ): VirtualCellValueNode {
    let nodeKey = cellAddressToKey(cellAddress).replace(/^[^:]+:/, "virtual:");
    nodeKey += ":";
    const normalizedCellValue = normalizeSerializedCellValue(cellValue);

    if (
      typeof normalizedCellValue === "string" &&
      normalizedCellValue.startsWith("=")
    ) {
      const ast = parseFormula(normalizedCellValue.slice(1));
      nodeKey += "ast:" + astToString(ast);
    } else if (
      typeof normalizedCellValue === "string" &&
      normalizedCellValue !== ""
    ) {
      nodeKey += "string:" + normalizedCellValue;
    } else if (typeof normalizedCellValue === "number") {
      nodeKey += "number:" + normalizedCellValue;
    } else if (typeof normalizedCellValue === "boolean") {
      nodeKey += "boolean:" + normalizedCellValue;
    } else if (normalizedCellValue === undefined) {
      nodeKey += "string:";
    } else {
      throw new Error("Invalid cell value: " + normalizedCellValue);
    }

    if (!this.virtualCellValueNodes.has(nodeKey)) {
      const node = new VirtualCellValueNode(nodeKey, cellAddress, cellValue);
      this.virtualCellValueNodes.set(nodeKey, node);
      return node;
    }
    return this.virtualCellValueNodes.get(nodeKey)!;
  }

  getRangeNode(rangeKey: string): RangeEvaluationNode {
    if (!rangeKey.startsWith("range:")) {
      throw new Error("Invalid range node key: " + rangeKey);
    }
    if (!this.ranges.has(rangeKey)) {
      const node = new RangeEvaluationNode(
        rangeKey,
        this,
        this.workbookManager
      );
      this.ranges.set(rangeKey, node);
      return node;
    }
    return this.ranges.get(rangeKey)!;
  }

  asts: Map<
    /**
     * ast key
     */
    string,
    {
      entries: Map<
        /**
         * context dependency key
         */
        string,
        {
          evalNode: AstEvaluationNode;
          contextDependency: ContextDependency;
        }
      >;
    }
  > = new Map();

  getAstNode(
    ast: ASTNode,
    currentContext: Omit<Required<ContextDependency>, "tableName"> & {
      tableName?: string;
    }
  ): AstEvaluationNode {
    const astKey = `ast:${astToString(ast)}`; // cache normalize this later
    const astEntries = this.asts.get(astKey);

    // if any of the ast entries match the current context, then we can return the ast node
    // otherwise we have to evalute the ast node to understand if it is context dependent
    // and later it will be saved using saveAstNode
    // eligibleKeysForContext returns keys ordered from most-specific to least-specific
    const keys = eligibleKeysForContext(currentContext);

    for (const key of keys) {
      const astEntry = astEntries?.entries.get(key);
      if (!astEntry) {
        continue;
      }

      if (astEntry) {
        return astEntry.evalNode;
      }
    }

    // by default the ast node is cell specific
    // but later, setContextDependency is called with a more open context dependency
    const node = new AstEvaluationNode(ast, currentContext);
    // initially we store it as a cell, sheet, workbook and table dependent node
    // but later, once resolved, we can store it under a looser dependency key, e.g. only sheet, workbook and table dependent
    this.saveAstNode(node, currentContext);
    return node;
  }


  /**
   * Once an AST node is evaluated, we know if it is context dependent
   * and will thus save it under the correct cache key according to its
   * contextDependency
   *
   * only resolved ast nodes can be saved
   */
  private saveAstNode(
    ast: AstEvaluationNode,
    contextDependency: ContextDependency
  ) {
    const astKey = ast.key;
    const contextDependencyKey = getContextDependencyKey(contextDependency);
    this.removeAstNodeFromCache(ast);
    const astEntries = this.asts.get(astKey);

    if (astEntries) {
      // if we don't already have an entry, then let's add it
      astEntries.entries.set(contextDependencyKey, {
        evalNode: ast,
        contextDependency,
      });
    } else {
      this.asts.set(astKey, {
        entries: new Map([
          [contextDependencyKey, { evalNode: ast, contextDependency }],
        ]),
      });
    }
  }

  private isFiniteEndAfter(value: number, end: RangeAddress["range"]["end"]["row"]) {
    return end.type === "infinity" || value < end.value;
  }

  private doesAddressAffectWatcherFrontier(
    address: CellAddress,
    watcher: FrontierWatcherNode
  ): boolean {
    const watcherRange = this.getWatcherRange(watcher).range;
    const rowWithinRange =
      address.rowIndex >= watcherRange.start.row &&
      (watcherRange.end.row.type === "infinity" ||
        address.rowIndex <= watcherRange.end.row.value);
    const colWithinRange =
      address.colIndex >= watcherRange.start.col &&
      (watcherRange.end.col.type === "infinity" ||
        address.colIndex <= watcherRange.end.col.value);
    const canReachFurtherRows = this.isFiniteEndAfter(
      address.rowIndex,
      watcherRange.end.row
    );
    const canReachFurtherCols = this.isFiniteEndAfter(
      address.colIndex,
      watcherRange.end.col
    );

    return (
      (rowWithinRange && canReachFurtherCols) ||
      (colWithinRange && canReachFurtherRows) ||
      (canReachFurtherRows && canReachFurtherCols)
    );
  }

  private collectCoverageWatchers(address: CellAddress): Set<FrontierWatcherNode> {
    const watchers =
      this.coverageWatchersBySheet.get(this.getSheetWatcherKey(address)) ??
      new Set<FrontierWatcherNode>();
    const affected = new Set<FrontierWatcherNode>();

    for (const watcher of watchers) {
      if (isCellInRange(address, this.getWatcherRange(watcher).range)) {
        affected.add(watcher);
      }
    }

    return affected;
  }

  private collectFrontierWatchers(address: CellAddress): Set<FrontierWatcherNode> {
    const watchers =
      this.frontierWatchersBySheet.get(this.getSheetWatcherKey(address)) ??
      new Set<FrontierWatcherNode>();
    const affected = new Set<FrontierWatcherNode>();

    for (const watcher of watchers) {
      if (this.doesAddressAffectWatcherFrontier(address, watcher)) {
        affected.add(watcher);
      }
    }

    return affected;
  }

  private collectWatchersIntersectingRange(address: RangeAddress): Set<FrontierWatcherNode> {
    const watchers =
      this.coverageWatchersBySheet.get(this.getSheetWatcherKey(address)) ??
      new Set<FrontierWatcherNode>();
    const affected = new Set<FrontierWatcherNode>();

    for (const watcher of watchers) {
      if (
        checkRangeIntersection(
          address.range,
          this.getWatcherRange(watcher).range
        )
      ) {
        affected.add(watcher);
      }
    }

    return affected;
  }

  private getNodeDependents(node: DependencyNode): Set<DependencyNode> {
    return (this.hardDependents.get(node) ?? new Set()).union(
      this.frontierDependents.get(node) ?? new Set()
    );
  }

  private collectInvalidationExtras(node: DependencyNode): Set<DependencyNode> {
    const extras = new Set<DependencyNode>();

    if (node instanceof CellValueNode || node instanceof SpillMetaNode) {
      const spillMetaNode = this.getLinkedSpillMetaNode(node);
      if (spillMetaNode) {
        extras.add(spillMetaNode);
      }

      const cellNode = this.getLinkedCellValueNode(node);
      if (cellNode) {
        extras.add(cellNode);
      }

      const spill = this.getSpilledValue(node.key);
      if (spill) {
        for (const watcher of this.collectWatchersIntersectingRange({
          workbookName: spill.origin.workbookName,
          sheetName: spill.origin.sheetName,
          range: spill.spillOnto,
        })) {
          extras.add(watcher);
        }
      }

      const cellAddress = node.cellAddress;
      for (const watcher of this.collectCoverageWatchers(cellAddress)) {
        extras.add(watcher);
      }
      for (const watcher of this.collectFrontierWatchers(cellAddress)) {
        extras.add(watcher);
      }
    }

    return extras;
  }

  private invalidateNodeState(node: DependencyNode, invalidatedKeys: Set<string>) {
    this.unregisterNode(node);
    invalidatedKeys.add(node.key);

    if (node instanceof CellValueNode) {
      node.clearSpillMetaNode();
      this.deleteSpilledValue(node.key);
      node.invalidate();
      return;
    }

    if (node instanceof SpillMetaNode) {
      this.deleteSpilledValue(node.key);
      this.getLinkedCellValueNode(node)?.clearSpillMetaNode();
      node.invalidate();
      return;
    }

    if (
      node instanceof EmptyCellEvaluationNode ||
      node instanceof RangeEvaluationNode ||
      node instanceof AstEvaluationNode
    ) {
      node.invalidate();
      return;
    }

    if (node instanceof ResourceDependencyNode) {
      return;
    }
  }

  private isResourceNodeInRemovedScope(
    resourceKey: string,
    scope: RemovedScope
  ): boolean {
    if (scope.type === "workbook") {
      return (
        resourceKey === `resource:workbook:${scope.workbookName}` ||
        resourceKey.startsWith(`resource:sheet:${scope.workbookName}:`) ||
        resourceKey.startsWith(`resource:table:${scope.workbookName}:`) ||
        resourceKey.startsWith(
          `resource:named:workbook:${scope.workbookName}:`
        ) ||
        resourceKey.startsWith(`resource:named:sheet:${scope.workbookName}:`)
      );
    }

    return (
      resourceKey ===
        `resource:sheet:${scope.workbookName}:${scope.sheetName}` ||
      resourceKey.startsWith(
        `resource:named:sheet:${scope.workbookName}:${scope.sheetName}:`
      )
    );
  }

  private isNodeInRemovedScope(node: DependencyNode, scope: RemovedScope): boolean {
    if (node instanceof ResourceDependencyNode) {
      return this.isResourceNodeInRemovedScope(node.key, scope);
    }

    if (node instanceof AstEvaluationNode) {
      const contextDependency = node.getContextDependency();
      if (scope.type === "workbook") {
        return contextDependency.workbookName === scope.workbookName;
      }
      return (
        contextDependency.workbookName === scope.workbookName &&
        contextDependency.sheetName === scope.sheetName
      );
    }

    const workbookName =
      node instanceof RangeEvaluationNode
        ? node.address.workbookName
        : node.cellAddress.workbookName;
    const sheetName =
      node instanceof RangeEvaluationNode
        ? node.address.sheetName
        : node.cellAddress.sheetName;

    if (scope.type === "workbook") {
      return workbookName === scope.workbookName;
    }

    return workbookName === scope.workbookName && sheetName === scope.sheetName;
  }

  private collectNodesForRemovedScopes(
    removedScopes: RemovedScope[]
  ): Set<DependencyNode> {
    const affected = new Set<DependencyNode>();

    for (const node of this.getAllPersistentNodes()) {
      if (removedScopes.some((scope) => this.isNodeInRemovedScope(node, scope))) {
        affected.add(node);
      }
    }

    return affected;
  }

  private purgeRemovedScopes(removedScopes: RemovedScope[]) {
    const shouldRemoveAddress = (
      address: Pick<CellAddress, "workbookName" | "sheetName">
    ) =>
      removedScopes.some((scope) =>
        scope.type === "workbook"
          ? address.workbookName === scope.workbookName
          : address.workbookName === scope.workbookName &&
            address.sheetName === scope.sheetName
      );

    for (const [key, node] of Array.from(this.cellNodes.entries())) {
      if (shouldRemoveAddress(node.cellAddress)) {
        this.unregisterNode(node);
        this.cellNodes.delete(key);
      }
    }

    for (const [key, node] of Array.from(this.spillMetaNodes.entries())) {
      if (shouldRemoveAddress(node.cellAddress)) {
        this.unregisterNode(node);
        this.spillMetaNodes.delete(key);
      }
    }

    for (const [key, node] of Array.from(this.emptyCells.entries())) {
      if (shouldRemoveAddress(node.cellAddress)) {
        this.unregisterNode(node);
        this.emptyCells.delete(key);
      }
    }

    for (const [key, node] of Array.from(this.ranges.entries())) {
      if (
        shouldRemoveAddress({
          workbookName: node.address.workbookName,
          sheetName: node.address.sheetName,
        })
      ) {
        this.unregisterNode(node);
        this.ranges.delete(key);
      }
    }

    for (const [key, node] of Array.from(this.resourceNodes.entries())) {
      if (removedScopes.some((scope) => this.isResourceNodeInRemovedScope(key, scope))) {
        this.unregisterNode(node);
        this.resourceNodes.delete(key);
      }
    }

    for (const [astKey, astEntries] of Array.from(this.asts.entries())) {
      for (const [contextKey, astEntry] of Array.from(astEntries.entries.entries())) {
        if (
          removedScopes.some((scope) =>
            this.isNodeInRemovedScope(astEntry.evalNode, scope)
          )
        ) {
          this.unregisterNode(astEntry.evalNode);
          astEntries.entries.delete(contextKey);
        }
      }

      if (astEntries.entries.size === 0) {
        this.asts.delete(astKey);
      }
    }

    for (const [spillOriginKey, spilledValue] of Array.from(this._spilledValues.entries())) {
      if (
        shouldRemoveAddress({
          workbookName: spilledValue.origin.workbookName,
          sheetName: spilledValue.origin.sheetName,
        })
      ) {
        this._spilledValues.delete(spillOriginKey);
      }
    }
  }

  public invalidateFromMutation(footprint: MutationInvalidation): void {
    const queue: DependencyNode[] = [];
    const visited = new Set<DependencyNode>();
    const invalidatedNodeKeys = new Set<string>();
    const invalidatedAstNodes = new Set<AstEvaluationNode>();
    const evictedAstNodes = new Set<AstEvaluationNode>();

    const evictAstNodeFromCache = (astNode: AstEvaluationNode) => {
      if (evictedAstNodes.has(astNode)) {
        return;
      }
      evictedAstNodes.add(astNode);
      this.removeAstNodeFromCache(astNode);
      invalidatedNodeKeys.add(astNode.key);
    };

    const invalidateAstNode = (astNode: AstEvaluationNode) => {
      if (invalidatedAstNodes.has(astNode)) {
        return;
      }
      invalidatedAstNodes.add(astNode);
      evictAstNodeFromCache(astNode);
      this.unregisterNode(astNode);
      astNode.invalidate();
    };

    const tableContextChangedCells = Array.from(
      new Map(
        (footprint.tableContextChangedCells ?? []).map((address) => [
          cellAddressToKey(address),
          address,
        ])
      ).values()
    );

    for (const astNode of this.collectExistingAstNodesForCells(
      tableContextChangedCells
    )) {
      evictAstNodeFromCache(astNode);
    }

    for (const astNode of this.collectOrphanedOldFormulaAstNodesForCells(
      tableContextChangedCells
    )) {
      invalidateAstNode(astNode);
    }

    for (const touchedCell of footprint.touchedCells) {
      if (touchedCell.beforeKind === "formula") {
        for (const astNode of this.collectOrphanedOldFormulaAstNodesForCell(
          touchedCell.address
        )) {
          invalidateAstNode(astNode);
        }
      }

      for (const node of this.collectExistingNodesForCell(touchedCell.address)) {
        queue.push(node);
      }

      for (const node of this.collectSpillOriginsAffectingCell(touchedCell.address)) {
        queue.push(node);
      }

      for (const watcher of this.collectCoverageWatchers(touchedCell.address)) {
        queue.push(watcher);
      }
      for (const watcher of this.collectFrontierWatchers(touchedCell.address)) {
        queue.push(watcher);
      }
    }

    for (const resourceKey of footprint.resourceKeys) {
      const resourceNode = this.resourceNodes.get(resourceKey);
      if (resourceNode) {
        queue.push(resourceNode);
      }
    }

    if (footprint.removedScopes?.length) {
      for (const node of this.collectNodesForRemovedScopes(footprint.removedScopes)) {
        queue.push(node);
      }
    }

    while (queue.length > 0) {
      const node = queue.pop();
      if (!node || visited.has(node)) {
        continue;
      }
      visited.add(node);

      for (const dependent of this.getNodeDependents(node)) {
        queue.push(dependent);
      }

      for (const extra of this.collectInvalidationExtras(node)) {
        queue.push(extra);
      }

      this.invalidateNodeState(node, invalidatedNodeKeys);
    }

    this.cacheManager.deleteEvaluationOrders(invalidatedNodeKeys);
    this.cacheManager.clearSCCCache();

    if (footprint.removedScopes?.length) {
      this.purgeRemovedScopes(footprint.removedScopes);
    }
  }

  //#region dependency graph methods

  /**
   * Get transitive dependencies and transitive frontier dependencies
   * This is only used by buildEvaluationOrder, so we'll optimize it there
   */
  getTransitiveDepsForEvalOrder(
    node: DependencyNode,
    visited: Set<DependencyNode> = new Set()
  ): Set<DependencyNode> {
    // Prevent infinite recursion
    if (visited.has(node)) {
      return new Set();
    }

    // If the node is resolved, then we don't need to evaluate it
    if (node && node.resolved) {
      return new Set();
    }

    // Mark this node as visited for cycle detection
    visited.add(node);

    const allNodes = new Set<DependencyNode>();
    allNodes.add(node);

    // Get direct dependencies (regular + frontier)
    const directDeps = node.getDependencies();

    // Recursively get transitive dependencies for each direct dependency
    for (const dep of directDeps) {
      if (!visited.has(dep)) {
        const depTransitiveDeps = this.getTransitiveDepsForEvalOrder(
          dep,
          visited
        );
        for (const transitiveDep of depTransitiveDeps) {
          allNodes.add(transitiveDep);
        }
      }
    }

    // Remove this node from visited set for other branches
    visited.delete(node);

    return allNodes;
  }

  /**
   * Build evaluation order for a cell using SCC-based condensation DAG approach
   *
   * Algorithm:
   * 1. Discover all transitive dependencies (skipping resolved nodes)
   * 2. Find SCCs using Tarjan's algorithm
   * 3. Create condensation DAG from SCCs
   * 4. Topologically sort the condensation DAG using Kahn's algorithm
   * 5. For each SCC, create internal evaluation order with cycle breaking
   * 6. Join the sorted SCC evaluation orders to create final evaluation order
   */
  buildEvaluationOrder(
    node: CellValueNode | EmptyCellEvaluationNode | VirtualCellValueNode
  ): EvaluationOrder {
    if (node.resolved && this.cacheManager.getEvaluationOrder(node.key)) {
      return this.cacheManager.getEvaluationOrder(node.key)!;
    }

    // Phase 1: Discover all transitive dependencies (skipping resolved nodes)
    const allNodes = new Map<string, DependencyNode>();
    const visitedForDiscovery = new Set<DependencyNode>();

    const discoverNodes = (currentNode: DependencyNode) => {
      if (currentNode && currentNode.resolved) {
        return;
      }

      if (!allNodes.has(currentNode.key)) {
        allNodes.set(currentNode.key, currentNode);
      }

      if (visitedForDiscovery.has(currentNode)) {
        return;
      }

      visitedForDiscovery.add(currentNode);

      const allDeps = currentNode.getAllDependencies();
      for (const dep of allDeps) {
        discoverNodes(dep);
      }
    };

    discoverNodes(node);

    if (allNodes.size === 0 && node && node.resolved) {
      const result: EvaluationOrder = {
        evaluationOrder: new Set([node]),
        hasCycle: false,
        hash: this.computeHash(new Set([node])),
      };

      if (node && node.resolved) {
        this.cacheManager.setEvaluationOrder(node.key, result);
      }

      return result;
    }

    // Phase 2: Find SCCs using Tarjan's algorithm
    // Build SCCs considering ALL dependencies (soft + hard edges)
    const sccs = this.findSCCs(allNodes, true);

    // Phase 3: Create condensation DAG and check for cached SCCs
    const nodeToSCCId = new Map<DependencyNode, number>();
    const sccList: import("../types").SCC[] = [];

    for (let i = 0; i < sccs.length; i++) {
      const sccNodes = sccs[i]!;

      // Check if all nodes in this SCC are resolved
      const allResolved = Array.from(sccNodes).every((n) => n.resolved);

      // Create SCC hash for caching
      const sccHash = Array.from(sccNodes)
        .map((n) => n.key)
        .sort()
        .join("|");

      // Try to get cached SCC if it's resolved
      let scc: import("../types").SCC;
      const cachedSCC = allResolved
        ? this.cacheManager.getSCC(sccHash)
        : undefined;

      if (cachedSCC) {
        scc = cachedSCC;
      } else {
        // Build evaluation order for this SCC with cycle breaking
        const sccEvalOrder = this.buildSCCEvaluationOrder(sccNodes);

        // Find hard-edge SCCs within this soft-edge SCC
        // Hard-edge SCCs are formed by only regular dependencies
        const hardEdgeSCCs = this.findSCCs(
          new Map(Array.from(sccNodes).map((n) => [n.key, n])),
          false // Use only hard edges (regular dependencies)
        );

        scc = {
          id: i,
          nodes: sccNodes,
          evaluationOrder: sccEvalOrder,
          resolved: allResolved,
          hardEdgeSCCs,
        };

        // Cache if resolved
        if (allResolved) {
          this.cacheManager.setSCC(sccHash, scc);
        }
      }

      sccList.push(scc);

      for (const n of sccNodes) {
        nodeToSCCId.set(n, i);
      }
    }

    // Build SCC dependency graph
    // Edge from A to B means A depends on B, so B must be evaluated before A
    const sccGraph = new Map<number, Set<number>>();
    for (let i = 0; i < sccList.length; i++) {
      sccGraph.set(i, new Set());
    }

    for (const [_, n] of allNodes) {
      const nSCCId = nodeToSCCId.get(n)!;
      // Use ALL dependencies (regular + frontier) for the condensation DAG
      // This ensures proper evaluation order even with frontier dependencies
      const deps = n.getAllDependencies();

      for (const dep of deps) {
        if (!allNodes.has(dep.key)) continue;

        const depSCCId = nodeToSCCId.get(dep)!;
        // n depends on dep, so dep's SCC must come before n's SCC
        // Add edge from dep's SCC to n's SCC
        if (nSCCId !== depSCCId) {
          sccGraph.get(depSCCId)!.add(nSCCId);
        }
      }
    }

    // Phase 4: Topologically sort SCCs using Kahn's algorithm
    const inDegree = new Map<number, number>();
    for (let i = 0; i < sccList.length; i++) {
      inDegree.set(i, 0);
    }

    for (const [_, deps] of sccGraph) {
      for (const toId of deps) {
        inDegree.set(toId, inDegree.get(toId)! + 1);
      }
    }

    const queue: number[] = [];
    for (let i = 0; i < sccList.length; i++) {
      if (inDegree.get(i) === 0) {
        queue.push(i);
      }
    }

    const sortedSCCIds: number[] = [];
    while (queue.length > 0) {
      const sccId = queue.shift()!;
      sortedSCCIds.push(sccId);

      const deps = sccGraph.get(sccId)!;
      for (const depId of deps) {
        const newInDegree = inDegree.get(depId)! - 1;
        inDegree.set(depId, newInDegree);
        if (newInDegree === 0) {
          queue.push(depId);
        }
      }
    }

    // Phase 5: Join evaluation orders from sorted SCCs
    const evaluationOrderArray: DependencyNode[] = [];
    for (const sccId of sortedSCCIds) {
      const scc = sccList[sccId]!;
      evaluationOrderArray.push(...scc.evaluationOrder);
    }

    const evaluationOrder = new Set(evaluationOrderArray);

    // Identify cycle nodes from hard-edge SCCs
    const cycleNodes = new Set<DependencyNode>();
    for (const scc of sccList) {
      for (const hardEdgeSCC of scc.hardEdgeSCCs) {
        // A hard-edge SCC with multiple nodes or a self-loop indicates a real cycle
        if (hardEdgeSCC.size > 1) {
          for (const n of hardEdgeSCC) {
            cycleNodes.add(n);
          }
        } else if (hardEdgeSCC.size === 1) {
          const node = Array.from(hardEdgeSCC)[0]!;
          if (node.getDependencies().has(node)) {
            cycleNodes.add(node);
          }
        }
      }
    }

    const hasCycle = cycleNodes.size > 0;
    const result: EvaluationOrder = {
      evaluationOrder,
      hasCycle,
      ...(hasCycle && { cycleNodes }),
      hash: this.computeGraphHash(allNodes, sccList),
      sccDAG: {
        sccList,
        sccGraph,
      },
    };

    if (node && node.resolved) {
      this.cacheManager.setEvaluationOrder(node.key, result);
    }

    return result;
  }

  /**
   * Find strongly connected components using Tarjan's algorithm
   * @param nodes - Map of nodes to analyze
   * @param includeFrontier - If true, use getAllDependencies(); if false, use getDependencies()
   * @returns Array of SCCs (each SCC is a Set of nodes)
   */
  private findSCCs(
    nodes: Map<string, DependencyNode>,
    includeFrontier: boolean
  ): Set<DependencyNode>[] {
    const index = new Map<DependencyNode, number>();
    const lowlink = new Map<DependencyNode, number>();
    const onStack = new Set<DependencyNode>();
    const stack: DependencyNode[] = [];
    const sccs: Set<DependencyNode>[] = [];
    let currentIndex = 0;

    const strongConnect = (v: DependencyNode) => {
      index.set(v, currentIndex);
      lowlink.set(v, currentIndex);
      currentIndex++;
      stack.push(v);
      onStack.add(v);

      // Use either all dependencies or just regular dependencies
      const successors = includeFrontier
        ? v.getAllDependencies()
        : v.getDependencies();

      for (const w of successors) {
        if (!nodes.has(w.key)) {
          continue;
        }

        if (!index.has(w)) {
          strongConnect(w);
          lowlink.set(v, Math.min(lowlink.get(v)!, lowlink.get(w)!));
        } else if (onStack.has(w)) {
          lowlink.set(v, Math.min(lowlink.get(v)!, index.get(w)!));
        }
      }

      if (lowlink.get(v) === index.get(v)) {
        const scc = new Set<DependencyNode>();
        let w: DependencyNode;
        do {
          w = stack.pop()!;
          onStack.delete(w);
          scc.add(w);
        } while (w !== v);

        sccs.push(scc);
      }
    };

    for (const [_, n] of nodes) {
      if (!index.has(n)) {
        strongConnect(n);
      }
    }

    return sccs;
  }

  /**
   * Build evaluation order within a single SCC using DFS with cycle breaking
   * Uses all dependencies (including frontier) for proper evaluation ordering
   */
  private buildSCCEvaluationOrder(
    sccNodes: Set<DependencyNode>
  ): DependencyNode[] {
    const visited = new Set<DependencyNode>();
    const visiting = new Set<DependencyNode>();
    const result: DependencyNode[] = [];

    const dfs = (n: DependencyNode) => {
      if (visited.has(n)) {
        return;
      }

      if (visiting.has(n)) {
        // Cycle detected (from any edge type), break it
        return;
      }

      visiting.add(n);

      // Use all dependencies for evaluation ordering (regular + frontier)
      const deps = n.getAllDependencies();
      for (const dep of deps) {
        if (sccNodes.has(dep) && !visited.has(dep)) {
          dfs(dep);
        }
      }

      visiting.delete(n);
      visited.add(n);
      result.push(n);
    };

    // Sort nodes by key for deterministic ordering
    const sortedNodes = Array.from(sccNodes).sort((a, b) =>
      a.key.localeCompare(b.key)
    );

    for (const n of sortedNodes) {
      if (!visited.has(n)) {
        dfs(n);
      }
    }

    return result;
  }

  /**
   * Compute hash representing the graph structure including SCC information
   */
  private computeGraphHash(
    allNodes: Map<string, DependencyNode>,
    sccList: import("../types").SCC[]
  ): string {
    const parts: string[] = [];

    // Hash nodes and their dependencies
    for (const [key, node] of Array.from(allNodes.entries()).sort()) {
      const deps = Array.from(node.getAllDependencies())
        .map((d) => d.key)
        .sort()
        .join(",");
      parts.push(`${key}:[${deps}]`);
    }

    // Add SCC structure
    for (const scc of sccList) {
      const nodeKeys = Array.from(scc.nodes)
        .map((n) => n.key)
        .sort()
        .join(",");
      parts.push(`SCC${scc.id}:{${nodeKeys}}`);
    }

    return parts.join("|");
  }

  /**
   * Compute a hash representing the current state of evaluated nodes
   * This hash changes when dependencies, frontier dependencies, or discarded frontier dependencies change
   */
  private computeHash(allNodes: Set<DependencyNode>): string {
    const nodeStates: string[] = [];

    for (const node of Array.from(allNodes).sort()) {
      if (node) {
        const deps = Array.from(node.getDependencies() || [])
          .map((dep) => dep.key)
          .sort()
          .join(",");

        // Handle frontier dependencies (Map<string, Set<string>>)
        const frontierDeps: string = Array.from(node.getFrontierDependencies())
          .map((dep) => dep.key)
          .sort()
          .join(";");

        const nodeState = `${node.key}:{deps:[${deps}],frontier:[${frontierDeps}]}`;
        nodeStates.push(nodeState);
      }
    }

    return nodeStates.join("|");
  }

  /**
   * Get a hierarchical dependency tree for a node
   */
  getDependencyTree(node: DependencyNode): DependencyTreeNode {
    const visited = new Set<DependencyNode>();

    const nodeToType = (node: DependencyNode): "cell" | "range" | "empty" => {
      if (node instanceof RangeEvaluationNode) {
        return "range";
      }
      if (node instanceof EmptyCellEvaluationNode) {
        return "empty";
      }
      return "cell";
    };

    const buildTree = (
      node: DependencyNode,
      isSelf = false
    ): DependencyTreeNode => {
      const cellRef: string = node.toString();

      // Handle self-reference to avoid infinite recursion
      if (isSelf) {
        return {
          type: nodeToType(node),
          resultType:
            node instanceof RangeEvaluationNode
              ? "range"
              : node.evaluationResult
              ? node.evaluationResult.type
              : "awaiting-evaluation",
          canResolve: node.canResolve(),
          key: cellRef,
          directDepsUpdated: node.directDepsUpdated,
          resolved: node.resolved,
          self: true,
          circular: true,
        };
      }

      // Avoid infinite recursion for circular dependencies
      if (visited.has(node)) {
        return {
          type: nodeToType(node),
          resultType:
            node instanceof RangeEvaluationNode
              ? "range"
              : node.evaluationResult
              ? node.evaluationResult.type
              : "awaiting-evaluation",
          canResolve: node.canResolve(),
          key: cellRef,
          directDepsUpdated: node.directDepsUpdated,
          resolved: node.resolved,
          circular: true,
        };
      }

      visited.add(node);

      const directDeps = Array.from(node.getDependencies());
      let frontierDeps = Array.from(node.getFrontierDependencies());

      // Get regular dependencies
      const deps: DependencyTreeNode[] = directDeps.map((dep) =>
        buildTree(dep, dep.key === node.key)
      );

      const frontierDependencies: DependencyTreeNode[] = frontierDeps.map(
        (dep) => buildTree(dep, false)
      );

      visited.delete(node);

      const result: DependencyTreeNode = {
        type: nodeToType(node),
        resultType:
          node instanceof RangeEvaluationNode
            ? "range"
            : node.evaluationResult
            ? node.evaluationResult.type
            : "awaiting-evaluation",
        canResolve: node.canResolve(),
        key: cellRef,
        directDepsUpdated: node.directDepsUpdated,
        resolved: node.resolved,
      };

      // Only include deps and frontierDependencies if they have content
      if (deps.length > 0) {
        result.deps = deps;
      }
      if (frontierDependencies.length > 0) {
        result.frontierDependencies = frontierDependencies;
      }

      return result;
    };

    return buildTree(node);
  }
  //#endregion

  markResolvedNodes(node: DependencyNode): void {
    // Track visited nodes to avoid infinite loops in circular dependencies
    const visited = new Set<DependencyNode>();
    visited.add(node); // Don't revisit the current cell

    const areTransitiveDepsResolved = (nodes: Set<DependencyNode>): boolean => {
      let canResolve = true;
      for (const node of nodes) {
        if (visited.has(node) || node.resolved) {
          continue;
        }
        visited.add(node);

        // Check the node's dependencies to not cause cycles with frontier dependencies
        const directDeps = node.getDependencies();

        const a = areTransitiveDepsResolved(directDeps);
        const b = node.canResolve();

        if (!a || !b) {
          canResolve = false;
        }
        if (a && b) {
          node.resolve();
          // if an ast node is resolved, it will get removed from the dependency graph
          // and thus never reach evaluateNode in formula evaluator
          // so we need to save it here. The latest context dependency is the correct one.
          if (node instanceof AstEvaluationNode) {
            this.saveAstNode(node, node.getContextDependency());
          }
        }
      }
      return canResolve;
    };

    if (
      areTransitiveDepsResolved(node.getDependencies()) &&
      node.canResolve()
    ) {
      node.resolve();
      if (node instanceof AstEvaluationNode) {
        this.saveAstNode(node, node.getContextDependency());
      }
    }
  }

  /**
   * Update SCCs in cache to mark them as resolved if all their nodes are resolved
   */
  public updateResolvedSCCs(evalOrder: EvaluationOrder): void {
    if (!evalOrder.sccDAG) {
      return;
    }

    // Check each SCC and update cache if all nodes are resolved
    for (const scc of evalOrder.sccDAG.sccList) {
      const allResolved = Array.from(scc.nodes).every((n) => n.resolved);

      if (allResolved && !scc.resolved) {
        // Create updated SCC with resolved flag
        const updatedSCC: import("../types").SCC = {
          ...scc,
          resolved: true,
        };

        // Update cache
        const sccHash = Array.from(scc.nodes)
          .map((n) => n.key)
          .sort()
          .join("|");

        this.cacheManager.setSCC(sccHash, updatedSCC);
      }
    }
  }

}
