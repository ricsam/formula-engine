import type { EvaluationOrder } from "../types";

export class CacheManager {
  private _evaluationOrderCache = new Map<string, EvaluationOrder>();
  constructor() {}

  setEvaluationOrder(nodeKey: string, evaluationOrder: EvaluationOrder): void {
    this._evaluationOrderCache.set(nodeKey, evaluationOrder);
  }

  getEvaluationOrder(nodeKey: string): EvaluationOrder | undefined {
    return this._evaluationOrderCache.get(nodeKey);
  }
}
