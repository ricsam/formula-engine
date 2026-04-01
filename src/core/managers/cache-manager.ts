import type { EvaluationOrder, SCC } from "../types";
import type { DependencyNode } from "./dependency-node";
import type {
  CacheManagerSnapshot,
  NodeSnapshotId,
  SerializedSCCSnapshot,
} from "../engine-snapshot";

export class CacheManager {
  private _evaluationOrderCache = new Map<string, EvaluationOrder>();

  /**
   * Cache for SCCs - key is a hash of the node keys in the SCC
   */
  private _sccCache = new Map<string, SCC>();

  constructor() {}

  setEvaluationOrder(nodeKey: string, evaluationOrder: EvaluationOrder): void {
    this._evaluationOrderCache.set(nodeKey, evaluationOrder);
  }

  getEvaluationOrder(nodeKey: string): EvaluationOrder | undefined {
    return this._evaluationOrderCache.get(nodeKey);
  }

  deleteEvaluationOrder(nodeKey: string): void {
    this._evaluationOrderCache.delete(nodeKey);
  }

  deleteEvaluationOrders(nodeKeys: Iterable<string>): void {
    for (const nodeKey of nodeKeys) {
      this._evaluationOrderCache.delete(nodeKey);
    }
  }

  setSCC(sccHash: string, scc: SCC): void {
    this._sccCache.set(sccHash, scc);
  }

  getSCC(sccHash: string): SCC | undefined {
    return this._sccCache.get(sccHash);
  }

  clearSCCCache(): void {
    this._sccCache.clear();
  }

  clear(): void {
    this._evaluationOrderCache.clear();
    this._sccCache.clear();
  }

  toSnapshot(
    getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId | undefined
  ): CacheManagerSnapshot {
    const evaluationOrders: CacheManagerSnapshot["evaluationOrders"] = [];

    for (const [nodeKey, evaluationOrder] of this._evaluationOrderCache.entries()) {
      const orderedNodeIds = Array.from(evaluationOrder.evaluationOrder).map(
        getNodeSnapshotId
      );
      if (orderedNodeIds.some((nodeId) => nodeId === undefined)) {
        continue;
      }

      const cycleNodeIds = evaluationOrder.cycleNodes
        ? Array.from(evaluationOrder.cycleNodes).map(getNodeSnapshotId)
        : undefined;
      if (cycleNodeIds?.some((nodeId) => nodeId === undefined)) {
        continue;
      }

      evaluationOrders.push({
        nodeKey,
        evaluationOrder: orderedNodeIds as NodeSnapshotId[],
        hasCycle: evaluationOrder.hasCycle,
        cycleNodes: cycleNodeIds as NodeSnapshotId[] | undefined,
        hash: evaluationOrder.hash,
      });
    }

    const sccs: CacheManagerSnapshot["sccs"] = [];
    for (const [hash, scc] of this._sccCache.entries()) {
      const serializedSCC = this.serializeSCC(scc, getNodeSnapshotId);
      if (!serializedSCC) {
        continue;
      }
      sccs.push({ hash, scc: serializedSCC });
    }

    return {
      evaluationOrders,
      sccs,
    };
  }

  restoreFromSnapshot(
    snapshot: CacheManagerSnapshot,
    resolveNodeSnapshotId: (
      nodeId: NodeSnapshotId
    ) => DependencyNode | undefined
  ) {
    this.clear();

    for (const evaluationOrderSnapshot of snapshot.evaluationOrders) {
      const evaluationOrderNodes = evaluationOrderSnapshot.evaluationOrder
        .map(resolveNodeSnapshotId)
        .filter((node): node is DependencyNode => node !== undefined);
      if (evaluationOrderNodes.length !== evaluationOrderSnapshot.evaluationOrder.length) {
        continue;
      }

      const cycleNodes = evaluationOrderSnapshot.cycleNodes
        ?.map(resolveNodeSnapshotId)
        .filter((node): node is DependencyNode => node !== undefined);
      if (
        evaluationOrderSnapshot.cycleNodes &&
        cycleNodes &&
        cycleNodes.length !== evaluationOrderSnapshot.cycleNodes.length
      ) {
        continue;
      }

      this._evaluationOrderCache.set(evaluationOrderSnapshot.nodeKey, {
        evaluationOrder: new Set(evaluationOrderNodes),
        hasCycle: evaluationOrderSnapshot.hasCycle,
        cycleNodes: cycleNodes ? new Set(cycleNodes) : undefined,
        hash: evaluationOrderSnapshot.hash,
      });
    }

    for (const { hash, scc } of snapshot.sccs) {
      const deserialized = this.deserializeSCC(scc, resolveNodeSnapshotId);
      if (deserialized) {
        this._sccCache.set(hash, deserialized);
      }
    }
  }

  private serializeSCC(
    scc: SCC,
    getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId | undefined
  ): SerializedSCCSnapshot | undefined {
    const nodeIds = Array.from(scc.nodes).map(getNodeSnapshotId);
    if (nodeIds.some((nodeId) => nodeId === undefined)) {
      return undefined;
    }

    const evaluationOrder = scc.evaluationOrder.map(getNodeSnapshotId);
    if (evaluationOrder.some((nodeId) => nodeId === undefined)) {
      return undefined;
    }

    const hardEdgeSCCs = scc.hardEdgeSCCs.map((group) =>
      Array.from(group).map(getNodeSnapshotId)
    );
    if (
      hardEdgeSCCs.some((group) =>
        group.some((nodeId) => nodeId === undefined)
      )
    ) {
      return undefined;
    }

    return {
      id: scc.id,
      nodes: nodeIds as NodeSnapshotId[],
      evaluationOrder: evaluationOrder as NodeSnapshotId[],
      resolved: scc.resolved,
      hardEdgeSCCs: hardEdgeSCCs as NodeSnapshotId[][],
    };
  }

  private deserializeSCC(
    snapshot: SerializedSCCSnapshot,
    resolveNodeSnapshotId: (
      nodeId: NodeSnapshotId
    ) => DependencyNode | undefined
  ): SCC | undefined {
    const nodes = snapshot.nodes
      .map(resolveNodeSnapshotId)
      .filter((node): node is DependencyNode => node !== undefined);
    const evaluationOrder = snapshot.evaluationOrder
      .map(resolveNodeSnapshotId)
      .filter((node): node is DependencyNode => node !== undefined);
    const hardEdgeSCCs = snapshot.hardEdgeSCCs.map((group) =>
      group
        .map(resolveNodeSnapshotId)
        .filter((node): node is DependencyNode => node !== undefined)
    );

    if (
      nodes.length !== snapshot.nodes.length ||
      evaluationOrder.length !== snapshot.evaluationOrder.length ||
      hardEdgeSCCs.some(
        (group, index) => group.length !== snapshot.hardEdgeSCCs[index]?.length
      )
    ) {
      return undefined;
    }

    return {
      id: snapshot.id,
      nodes: new Set(nodes),
      evaluationOrder,
      resolved: snapshot.resolved,
      hardEdgeSCCs: hardEdgeSCCs.map((group) => new Set(group)),
    };
  }
}
