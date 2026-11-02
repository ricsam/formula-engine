import type { ContextDependency } from "../evaluator/evaluation-context";
import type { DependencyNode } from "./managers/dependency-node";
import type {
  CellAddress,
  CellInRangeResult,
  CellValue,
  ConditionalStyle,
  DirectCellDataType,
  DirectCellStyle,
  FormulaError,
  NamedExpression,
  RangeAddress,
  RangeMetadata,
  RelativeRange,
  SpreadsheetRange,
  TableDefinition,
  TrackedReference,
  Workbook,
} from "./types";

export const ENGINE_SNAPSHOT_VERSION = 6 as const;
export const LEGACY_ENGINE_SNAPSHOT_VERSION = 5 as const;

export type NodeSnapshotId = string;

export type NamedExpressionManagerSnapshot = {
  sheetExpressions: Map<string, Map<string, Map<string, NamedExpression>>>;
  workbookExpressions: Map<string, Map<string, NamedExpression>>;
  globalExpressions: Map<string, NamedExpression>;
};

export type WorkbookManagerSnapshot = Map<string, Workbook>;

export type TableManagerSnapshot = Map<string, Map<string, TableDefinition>>;

export type StyleManagerSnapshot = {
  conditionalStyles: ConditionalStyle[];
  cellStyles: DirectCellStyle[];
  cellDataTypes: DirectCellDataType[];
};

export type RangeMetadataManagerSnapshot = RangeMetadata[];

export type ReferenceManagerSnapshot = Map<string, TrackedReference>;

export type EngineHistorySnapshotManagers = {
  workbook: WorkbookManagerSnapshot;
  namedExpression: NamedExpressionManagerSnapshot;
  table: TableManagerSnapshot;
  style: StyleManagerSnapshot;
  rangeMetadata: RangeMetadataManagerSnapshot;
  reference: ReferenceManagerSnapshot;
};

export type EngineHistorySnapshot = {
  version: typeof ENGINE_SNAPSHOT_VERSION;
  managers: EngineHistorySnapshotManagers;
};

export type SerializedValueEvaluationResultSnapshot = {
  type: "value";
  result: CellValue;
  sourceCell?: CellAddress;
};

export type SerializedErrorEvaluationResultSnapshot = {
  type: "error";
  err: FormulaError;
  message: string;
  errAddressId: NodeSnapshotId;
  sourceCell?: CellAddress;
};

export type SerializedSingleEvaluationResultSnapshot =
  | SerializedValueEvaluationResultSnapshot
  | SerializedErrorEvaluationResultSnapshot;

export type SerializedCellInRangeResultSnapshot = {
  relativePos: CellInRangeResult["relativePos"];
  result: SerializedSingleEvaluationResultSnapshot;
};

export type SerializedEvaluateAllCellsResultSnapshot =
  | SerializedErrorEvaluationResultSnapshot
  | {
      type: "values";
      values: SerializedCellInRangeResultSnapshot[];
    };

/**
 * A spilled result stored as fixed values: every cell read through the spill's
 * `evaluate` when the snapshot was taken. The values of a lazy spill such as
 * `A1:A2*1` come from cells its node never depended on, so the dependencies of
 * the node storing them also list the nodes those cells were read from.
 */
export type SerializedMaterializedSpillSnapshot = {
  kind: "materialized";
  relativeSpillArea: RelativeRange;
  source: string;
  sourceCell?: CellAddress;
  sourceRange?: RangeAddress;
  values: SerializedCellInRangeResultSnapshot[];
};

export type SerializedSourceRangeSpillSnapshot = {
  kind: "source-range";
  relativeSpillArea: RelativeRange;
  source: string;
  sourceCell?: CellAddress;
  sourceRange: RangeAddress;
};

export type SerializedSpillResultSnapshot =
  | SerializedMaterializedSpillSnapshot
  | SerializedSourceRangeSpillSnapshot;

export type SerializedSpilledValuesEvaluationResultSnapshot = {
  type: "spilled-values";
  spill: SerializedSpillResultSnapshot;
};

export type SerializedFunctionEvaluationResultSnapshot =
  | SerializedSingleEvaluationResultSnapshot
  | SerializedSpilledValuesEvaluationResultSnapshot;

export type SerializedSpillMetaEvaluationResultSnapshot =
  | SerializedErrorEvaluationResultSnapshot
  | SerializedSpilledValuesEvaluationResultSnapshot
  | {
      type: "does-not-spill";
    };

type SerializedBaseNodeSnapshot = {
  snapshotId: NodeSnapshotId;
  key: string;
  /**
   * The nodes this node's result was computed from. A restored node is trusted
   * as resolved and is only invalidated through these edges, so a snapshot
   * stores a node only together with every node it references (its
   * dependencies, its spill-meta node and the error addresses in its result);
   * a node that cannot be stored that way is left out, together with the nodes
   * that reference it, and recomputed on demand after a restore.
   */
  dependencies: NodeSnapshotId[];
};

export type SerializedCellValueNodeSnapshot = SerializedBaseNodeSnapshot & {
  kind: "cell-value";
  evaluationResult: SerializedSingleEvaluationResultSnapshot;
  spillMetaSnapshotId?: NodeSnapshotId;
};

export type SerializedSpillMetaNodeSnapshot = SerializedBaseNodeSnapshot & {
  kind: "spill-meta";
  evaluationResult: SerializedSpillMetaEvaluationResultSnapshot;
};

export type SerializedEmptyCellNodeSnapshot = SerializedBaseNodeSnapshot & {
  kind: "empty";
  evaluationResult: SerializedSingleEvaluationResultSnapshot;
};

export type SerializedRangeNodeSnapshot = SerializedBaseNodeSnapshot & {
  kind: "range";
  result: SerializedEvaluateAllCellsResultSnapshot;
};

export type SerializedAstNodeSnapshot = SerializedBaseNodeSnapshot & {
  kind: "ast";
  contextDependency: ContextDependency;
  evaluationResult: SerializedFunctionEvaluationResultSnapshot;
};

export type SerializedResourceNodeSnapshot = SerializedBaseNodeSnapshot & {
  kind: "resource";
};

export type SerializedDependencyNodeSnapshot =
  | SerializedCellValueNodeSnapshot
  | SerializedSpillMetaNodeSnapshot
  | SerializedEmptyCellNodeSnapshot
  | SerializedRangeNodeSnapshot
  | SerializedAstNodeSnapshot
  | SerializedResourceNodeSnapshot;

export type DependencyManagerSnapshot = {
  nodes: SerializedDependencyNodeSnapshot[];
  spilledValues: Array<[string, { origin: CellAddress; spillOnto: SpreadsheetRange }]>;
  /** Spills that evaluated to #SPILL!, with the area they would occupy. */
  blockedSpills: Array<[string, { origin: CellAddress; spillOnto: SpreadsheetRange }]>;
};

export type SerializedSCCSnapshot = {
  id: number;
  nodes: NodeSnapshotId[];
  evaluationOrder: NodeSnapshotId[];
  resolved: boolean;
  hardEdgeSCCs: NodeSnapshotId[][];
};

export type SerializedEvaluationOrderSnapshot = {
  nodeKey: string;
  evaluationOrder: NodeSnapshotId[];
  hasCycle: boolean;
  cycleNodes?: NodeSnapshotId[];
  hash: string;
};

export type CacheManagerSnapshot = {
  evaluationOrders: SerializedEvaluationOrderSnapshot[];
  sccs: Array<{
    hash: string;
    scc: SerializedSCCSnapshot;
  }>;
};

type EngineSnapshotManagers = EngineHistorySnapshotManagers & {
  dependency: DependencyManagerSnapshot;
  cache: CacheManagerSnapshot;
};

export type EngineSnapshot = {
  version: typeof ENGINE_SNAPSHOT_VERSION;
  managers: EngineSnapshotManagers;
};

export function getAstNodeSnapshotId(
  node: DependencyNode & { getContextDependency(): ContextDependency }
): NodeSnapshotId {
  return `${node.key}::${JSON.stringify(node.getContextDependency())}`;
}
