import type { EvaluationOrder, SCC } from "../types";

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

  setSCC(sccHash: string, scc: SCC): void {
    this._sccCache.set(sccHash, scc);
  }

  getSCC(sccHash: string): SCC | undefined {
    return this._sccCache.get(sccHash);
  }

  clear(): void {
    this._evaluationOrderCache.clear();
    this._sccCache.clear();
  }
}
