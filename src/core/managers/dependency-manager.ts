import {
  contextDependencyKeys,
  eligibleKeysForContext,
  getContextDependencyKey,
  type ContextDependency,
  type ContextDependencyType,
} from "src/evaluator/evaluation-context";
import {
  type CellAddress,
  type EvaluationOrder,
  type SpilledValue,
} from "../types";
import { isCellInRange, keyToCellAddress } from "../utils";

import { AstEvaluationNode } from "src/evaluator/dependency-nodes/ast-evaluation-node";
import { CellValueNode } from "src/evaluator/dependency-nodes/cell-value-node";
import { EmptyCellEvaluationNode } from "src/evaluator/dependency-nodes/empty-cell-evaluation-node";
import { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import type { ASTNode } from "src/parser/ast";
import { astToString } from "src/parser/formatter";
import { CacheManager } from "./cache-manager";
import type {
  CellNodeKeyDictionary,
  CellNodeType,
  DependencyNode,
} from "./dependency-node";
import { WorkbookManager } from "./workbook-manager";
import { SpillMetaNode } from "src/evaluator/dependency-nodes/spill-meta-node";
import { flags } from "src/debug/flags";

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
    | "does-not-spill";
  deps?: DependencyTreeNode[];
  frontierDependencies?: DependencyTreeNode[];
  self?: boolean;
  _debug?: {
    rawFrontierDependencies?: string[];
    discardedFrontierDependencies?: string[];
    activeFrontierDependencies?: string[];
  };
}

/**
 * The DependencyManager is responsible for storing the evaluated values and their dependencies.
 */
export class DependencyManager {
  /**
   * The dependency graph
   */
  private evaluatedNodes: Map<
    /**
     * key is the cell key, from cellAddressToKey
     */
    string,
    CellValueNode
  > = new Map();

  private spillMetaNodes: Map<string, SpillMetaNode> = new Map();

  private emptyCells: Map<string, EmptyCellEvaluationNode> = new Map();

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

  constructor(
    private cacheManager: CacheManager,
    private workbookManager: WorkbookManager
  ) {}

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
    this.evaluatedNodes.clear();
    this.emptyCells.clear();
    this.asts.clear();
    this.spillMetaNodes.clear();
    this.ranges.clear();
    this._spilledValues.clear();
  }

  setSpilledValue(nodeKey: string, spilledValue: SpilledValue): void {
    this._spilledValues.set(nodeKey.replace(/^[^:]+:/, ""), spilledValue);
  }

  getSpilledValue(nodeKey: string): SpilledValue | undefined {
    return this._spilledValues.get(nodeKey.replace(/^[^:]+:/, ""));
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
    if (!this.evaluatedNodes.has(nodeKey)) {
      const node = new CellValueNode(nodeKey);
      this.evaluatedNodes.set(nodeKey, node);
      return node;
    }
    return this.evaluatedNodes.get(nodeKey)!;
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

  lookupCellNode<T extends CellNodeType>(
    nodeKey: string,
    types: T[]
  ): CellNodeKeyDictionary[T][] {
    const cellAddress = keyToCellAddress(nodeKey);

    const emptyKey = nodeKey.replace(/^[^:]+:/, "empty:");
    const cellValueKey = nodeKey.replace(/^[^:]+:/, "cell-value:");
    const cellValueNode = this.evaluatedNodes.get(cellValueKey);
    const cellMetaKey = nodeKey.replace(/^[^:]+:/, "cell-meta:");
    const cellMetaNode = this.evaluatedNodes.get(cellMetaKey);

    if (this.workbookManager.isCellEmpty(cellAddress)) {
      if ((types as string[]).includes("empty")) {
        return [this.getEmptyCellNode(emptyKey) as any];
      } else {
        return [];
      }
    }

    const results: any[] = [];

    if ((types as string[]).includes("cell")) {
      if (cellValueNode) {
        results.push(cellValueNode);
      } else {
        results.push(this.getCellValueNode(cellValueKey));
      }
    }

    if (this.workbookManager.isFormulaCell(cellAddress)) {
      if ((types as string[]).includes("spill-meta")) {
        if (cellMetaNode) {
          results.push(cellMetaNode);
        } else {
          results.push(this.getSpillMetaNode(cellValueKey));
        }
      }
    }

    return results;
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

    const keys = eligibleKeysForContext(currentContext);

    for (const key of keys) {
      const astEntry = astEntries?.entries.get(key);
      if (astEntry) {
        return astEntry.evalNode;
      }
    }

    // if any of the ast entries match the current context, then we can return the ast node
    // otherwise we have to evalute the ast node to understand if it is context dependent
    // and later it will be saved using saveAstNode
    // if (astEntries) {
    //   for (const entry of astEntries.entries.values()) {
    //     // e.g. we have a row dependent ast dependency, dependent on row 1, and the current context dependency has rowIndex 1
    //     // then we will get a match
    //     const matches = Object.entries(entry.contextDependency).every(
    //       ([key, value]) => {
    //         if (typeof value === "undefined") {
    //           return true;
    //         }
    //         return currentContext[key as keyof ContextDependency] === value;
    //       }
    //     );
    //     if (matches) {
    //       return entry.evalNode;
    //     }
    //   }
    // }

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

  getEvaluatedNodes(): Map<string, CellValueNode> {
    return this.evaluatedNodes;
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
    node: CellValueNode | EmptyCellEvaluationNode
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
