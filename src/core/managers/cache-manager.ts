import type { CellAddress, EvaluationOrder } from "../types";
import type { RangeEvalOrderEntry } from "./range-eval-order-builder";

export class CacheManager {
  private _evaluationOrderCache = new Map<string, EvaluationOrder>();
  /**
   * Key is rangeKey + "@" + lookupOrder
   */
  private _rangeEvalOrderCache = new Map<string, RangeEvalOrderEntry[]>();

  /**
   * Key is rangeKey or cellKey
   */
  private _frontierCandidatesCache = new Map<string, CellAddress[]>();

  constructor() {}

  setEvaluationOrder(nodeKey: string, evaluationOrder: EvaluationOrder): void {
    this._evaluationOrderCache.set(nodeKey, evaluationOrder);
  }

  getEvaluationOrder(nodeKey: string): EvaluationOrder | undefined {
    return this._evaluationOrderCache.get(nodeKey);
  }

  setRangeEvalOrder(
    nodeKey: string,
    rangeEvalOrder: RangeEvalOrderEntry[]
  ): void {
    this._rangeEvalOrderCache.set(nodeKey, rangeEvalOrder);
  }

  getRangeEvalOrder(nodeKey: string): RangeEvalOrderEntry[] | undefined {
    return this._rangeEvalOrderCache.get(nodeKey);
  }

  setFrontierCandidates(
    nodeKey: string,
    frontierCandidates: CellAddress[]
  ): void {
    this._frontierCandidatesCache.set(nodeKey, frontierCandidates);
  }

  getFrontierCandidates(nodeKey: string): CellAddress[] | undefined {
    return this._frontierCandidatesCache.get(nodeKey);
  }

  clear(): void {
    this._evaluationOrderCache.clear();
    this._rangeEvalOrderCache.clear();
    this._frontierCandidatesCache.clear();
  }
}
