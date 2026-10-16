import { CellValueNode } from "./dependency-nodes/cell-value-node";
import type { RangeEvaluationNode } from "./range-evaluation-node";
import type { LookupOrder } from "../core/managers/range-eval-order-builder";
import type {
  EvaluateAllCellsResult,
  SingleEvaluationResult,
  SpilledValuesEvaluationResult,
} from "../core/types";
import type { EvaluationContext } from "./evaluation-context";
import type { FormulaEvaluator } from "./formula-evaluator";

type Values = Extract<EvaluateAllCellsResult, { type: "values" }>;
type PositionIndex = ReadonlyMap<string, SingleEvaluationResult>;

// Only raw range creation/restoration registers readers. sourceRange is metadata, not proof that
// a spill contains raw values (ROW/COLUMN and other transforms also carry it).
const readers = new WeakMap<SpilledValuesEvaluationResult, {
  evaluator: FormulaEvaluator;
  read: (context: EvaluationContext) => EvaluateAllCellsResult;
  canIndexDense?: (context: EvaluationContext) => boolean;
}>();
const indexes = new WeakMap<Values, PositionIndex>();
const denseSafety = new WeakMap<Values, boolean>();

export function canIndexDenseRangeNode(node: RangeEvaluationNode): boolean {
  const result = node.result;
  if (result.type !== "values") return false;
  const cached = denseSafety.get(result);
  if (cached !== undefined) return cached;
  // A resolved raw range with only ordinary cell dependencies has no spill
  // contributions. Be conservative for cycles and restored/other node kinds.
  const dependencies = node.getDependencies();
  const safe = result.values.length === dependencies.size &&
    [...dependencies].every(dependency =>
      dependency instanceof CellValueNode && !dependency.spillMeta
    );
  if (node.resolved) denseSafety.set(result, safe);
  return safe;
}

export function registerCriteriaRangeReader(
  range: SpilledValuesEvaluationResult,
  evaluator: FormulaEvaluator,
  read: (context: EvaluationContext) => EvaluateAllCellsResult,
  canIndexDense?: (context: EvaluationContext) => boolean
): SpilledValuesEvaluationResult {
  readers.set(range, { evaluator, read, canIndexDense });
  return range;
}

export function hasCriteriaRangeReader(
  range: SpilledValuesEvaluationResult,
  evaluator: FormulaEvaluator
): boolean {
  return readers.get(range)?.evaluator === evaluator;
}

// Dense consumers must retain cell callbacks when spill materialization may
// expose anchor-relative positions rather than coordinates of the raw range.
export function canIndexDenseCriteriaRange(
  range: SpilledValuesEvaluationResult,
  context: EvaluationContext
): boolean {
  return readers.get(range)?.canIndexDense?.(context) ?? false;
}

export function readCriteriaRange(
  range: SpilledValuesEvaluationResult,
  evaluator: FormulaEvaluator,
  context: EvaluationContext,
  lookupOrder: LookupOrder,
  needsIndex: boolean
): { result: EvaluateAllCellsResult; index?: PositionIndex } {
  const registered = readers.get(range);
  const trusted = registered?.evaluator === evaluator;
  // The reader always resolves the current range node and registers it with the
  // current consumer, even on an index hit. Never cache a node or its value here:
  // invalidation, undo and reset can replace either. Resolved range results are
  // replaced, not mutated, by RangeEvaluationNode.setResult/restore/invalidate.
  const result = trusted
    ? registered.read(context)
    : range.evaluateAllCells.call(evaluator, {
        context, evaluate: range.evaluate, origin: context.cellAddress, lookupOrder,
      });
  if (!needsIndex || result.type !== "values") return { result };

  let index = trusted ? indexes.get(result) : undefined;
  if (!index) {
    const map = new Map<string, SingleEvaluationResult>();
    for (const { result: cell, relativePos } of result.values) {
      map.set(`${relativePos.x},${relativePos.y}`, cell);
    }
    index = map;
    if (trusted) indexes.set(result, index);
  }
  return { result, index };
}
