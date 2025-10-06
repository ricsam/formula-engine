import type { CellAddress, EvaluationOrder } from "../types";

export class CacheManager {
  private _evaluationOrderCache = new Map<string, EvaluationOrder>();
  private _cellsInRangeCache = new Map<string, CellAddress[]>();
  private _frontierCandidatesCache = new Map<string, CellAddress[]>();
  constructor() {}

  setEvaluationOrder(nodeKey: string, evaluationOrder: EvaluationOrder): void {
    this._evaluationOrderCache.set(nodeKey, evaluationOrder);
  }

  getEvaluationOrder(nodeKey: string): EvaluationOrder | undefined {
    return this._evaluationOrderCache.get(nodeKey);
  }

  setCellsInRange(nodeKey: string, cellsInRange: CellAddress[]): void {
    this._cellsInRangeCache.set(nodeKey, cellsInRange);
  }

  getCellsInRange(nodeKey: string): CellAddress[] | undefined {
    return this._cellsInRangeCache.get(nodeKey);
  }

  setFrontierCandidates(nodeKey: string, frontierCandidates: CellAddress[]): void {
    this._frontierCandidatesCache.set(nodeKey, frontierCandidates);
  }

  getFrontierCandidates(nodeKey: string): CellAddress[] | undefined {
    return this._frontierCandidatesCache.get(nodeKey);
  }
}
