import { flags } from "../../debug/flags";
import { AstEvaluationNode } from "../../evaluator/dependency-nodes/ast-evaluation-node";
import { CellValueNode } from "../../evaluator/dependency-nodes/cell-value-node";
import { EvaluationContext } from "../../evaluator/evaluation-context";
import {
  EvaluationError,
  SheetNotFoundError,
} from "../../evaluator/evaluation-error";
import { RangeEvaluationNode } from "../../evaluator/range-evaluation-node";
import { normalizeSerializedCellValue } from "../../parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellInRangeResult,
  type CellValue,
  type ErrorEvaluationResult,
  type EvaluateAllCellsResult,
  type EvaluationOrder,
  type FunctionEvaluationResult,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type ValueEvaluationResult,
} from "../types";
import {
  captureEvaluationErrors,
  getAbsoluteRange,
  cellAddressToKey,
  checkRangeIntersection,
  getCellReference,
  getRangeIntersection,
  getRelativeRange,
  isCellInRange,
  parseCellReference,
  rangeAddressToKey,
} from "../utils";
import type { DependencyManager } from "./dependency-manager";
import type { WorkbookManager } from "./workbook-manager";
import { SpillMetaNode } from "../../evaluator/dependency-nodes/spill-meta-node";
import { EmptyCellEvaluationNode } from "../../evaluator/dependency-nodes/empty-cell-evaluation-node";
import type { TableManager } from "./table-manager";
import type { DependencyNode } from "./dependency-node";
import { VirtualCellValueNode } from "../../evaluator/dependency-nodes/virtual-cell-value-node";
import { ResourceDependencyNode } from "../../evaluator/dependency-nodes/resource-dependency-node";
import type {
  NodeSnapshotId,
  SerializedCellInRangeResultSnapshot,
  SerializedErrorEvaluationResultSnapshot,
  SerializedEvaluateAllCellsResultSnapshot,
  SerializedFunctionEvaluationResultSnapshot,
  SerializedSingleEvaluationResultSnapshot,
  SerializedSpillMetaEvaluationResultSnapshot,
  SerializedSpilledValuesEvaluationResultSnapshot,
} from "../engine-snapshot";
import type { MutationInvalidation } from "../commands/types";

export class EvaluationManager {
  private isEvaluating = false;

  constructor(
    private workbookManager: WorkbookManager,
    private tableManager: TableManager,
    private formulaEvaluator: FormulaEvaluator,
    private dependencyManager: DependencyManager
  ) {}

  clearEvaluationCache(): void {
    this.dependencyManager.clearEvaluationCache();
  }

  private dependsOnNode(
    candidate: DependencyNode,
    target: DependencyNode,
    visited: Set<DependencyNode> = new Set()
  ): boolean {
    if (candidate === target) {
      return true;
    }

    if (visited.has(candidate)) {
      return false;
    }
    visited.add(candidate);

    for (const dep of candidate.getDependencies()) {
      if (this.dependsOnNode(dep, target, visited)) {
        return true;
      }
    }

    return false;
  }

  invalidateFromMutation(footprint: MutationInvalidation): void {
    this.dependencyManager.invalidateFromMutation(footprint);
  }

  evaluationResultToSerializedValue(
    evaluation: SingleEvaluationResult,
    cellAddress: CellAddress,
    debug?: boolean
  ): SerializedCellValue {
    if (
      evaluation.type !== "error" &&
      evaluation.type !== "awaiting-evaluation"
    ) {
      const value = evaluation.result;

      return value.type === "infinity"
        ? value.sign === "positive"
          ? "INFINITY"
          : "-INFINITY"
        : value.value;
    }

    if (evaluation.type === "awaiting-evaluation") {
      return (
        evaluation.errAddress.key +
        " is awaiting evaluation of " +
        evaluation.waitingFor.key
      );
    }

    if (debug) {
      const errAddress = evaluation.errAddress.key;
      if (errAddress === cellAddressToKey(cellAddress)) {
        return evaluation.err + " " + evaluation.message;
      }
      return (
        evaluation.err +
        " in " +
        evaluation.errAddress.key +
        " " +
        evaluation.message
      );
    }

    return evaluation.err;
  }

  private buildSnapshotOrigin(
    node: CellValueNode | SpillMetaNode | AstEvaluationNode
  ): CellAddress {
    if ("cellAddress" in node) {
      return node.cellAddress;
    }

    const contextDependency = node.getContextDependency();
    return {
      workbookName: contextDependency.workbookName ?? "__snapshot__",
      sheetName: contextDependency.sheetName ?? "__snapshot__",
      colIndex: contextDependency.colIndex ?? 0,
      rowIndex: contextDependency.rowIndex ?? 0,
    };
  }

  private serializeErrorEvaluationResultSnapshot(
    evaluation: Exclude<ErrorEvaluationResult, { type: "awaiting-evaluation" }>,
    getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId
  ): SerializedErrorEvaluationResultSnapshot {
    return {
      type: "error",
      err: evaluation.err,
      message: evaluation.message,
      errAddressId: getNodeSnapshotId(evaluation.errAddress),
      sourceCell: evaluation.sourceCell,
    };
  }

  serializeSingleEvaluationResultSnapshot(
    evaluation: SingleEvaluationResult,
    getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId
  ): SerializedSingleEvaluationResultSnapshot | undefined {
    if (evaluation.type === "awaiting-evaluation") {
      return undefined;
    }

    if (evaluation.type === "error") {
      return this.serializeErrorEvaluationResultSnapshot(
        evaluation,
        getNodeSnapshotId
      );
    }

    return {
      type: "value",
      result: evaluation.result,
      sourceCell: evaluation.sourceCell,
    };
  }

  deserializeSingleEvaluationResultSnapshot(
    evaluation: SerializedSingleEvaluationResultSnapshot,
    resolveNodeSnapshotId: (nodeId: NodeSnapshotId) => DependencyNode
  ): SingleEvaluationResult {
    if (evaluation.type === "error") {
      return {
        type: "error",
        err: evaluation.err,
        message: evaluation.message,
        errAddress: resolveNodeSnapshotId(evaluation.errAddressId),
        sourceCell: evaluation.sourceCell,
      };
    }

    return {
      type: "value",
      result: evaluation.result,
      sourceCell: evaluation.sourceCell,
    };
  }

  serializeEvaluateAllCellsResultSnapshot(
    evaluation: EvaluateAllCellsResult,
    getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId
  ): SerializedEvaluateAllCellsResultSnapshot | undefined {
    if (evaluation.type === "awaiting-evaluation") {
      return undefined;
    }

    if (evaluation.type === "error") {
      return this.serializeErrorEvaluationResultSnapshot(
        evaluation,
        getNodeSnapshotId
      );
    }

    const values: SerializedCellInRangeResultSnapshot[] = [];
    for (const value of evaluation.values) {
      const serialized = this.serializeSingleEvaluationResultSnapshot(
        value.result,
        getNodeSnapshotId
      );
      if (!serialized) {
        return undefined;
      }
      values.push({
        relativePos: value.relativePos,
        result: serialized,
      });
    }

    return {
      type: "values",
      values,
    };
  }

  deserializeEvaluateAllCellsResultSnapshot(
    evaluation: SerializedEvaluateAllCellsResultSnapshot,
    resolveNodeSnapshotId: (nodeId: NodeSnapshotId) => DependencyNode
  ): EvaluateAllCellsResult {
    if (evaluation.type === "error") {
      return this.deserializeSingleEvaluationResultSnapshot(
        evaluation,
        resolveNodeSnapshotId
      ) as ErrorEvaluationResult;
    }

    return {
      type: "values",
      values: evaluation.values.map((value) => ({
        relativePos: value.relativePos,
        result: this.deserializeSingleEvaluationResultSnapshot(
          value.result,
          resolveNodeSnapshotId
        ),
      })),
    };
  }

  serializeFunctionEvaluationResultSnapshot(
    evaluation: FunctionEvaluationResult,
    options: {
      sourceNode: CellValueNode | SpillMetaNode | AstEvaluationNode;
      getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId;
    }
  ): SerializedFunctionEvaluationResultSnapshot | undefined {
    if (evaluation.type !== "spilled-values") {
      return this.serializeSingleEvaluationResultSnapshot(
        evaluation,
        options.getNodeSnapshotId
      );
    }

    const origin = this.buildSnapshotOrigin(options.sourceNode);
    const spillArea = evaluation.spillArea(origin);
    const relativeSpillArea = getRelativeRange(spillArea, {
      colIndex: origin.colIndex,
      rowIndex: origin.rowIndex,
    });

    if (evaluation.sourceRange) {
      return {
        type: "spilled-values",
        spill: {
          kind: "source-range",
          relativeSpillArea,
          source: evaluation.source,
          sourceCell: evaluation.sourceCell,
          sourceRange: evaluation.sourceRange,
        },
      };
    }

    if (
      relativeSpillArea.width.type !== "number" ||
      relativeSpillArea.height.type !== "number"
    ) {
      return undefined;
    }

    const context = new EvaluationContext(
      this.tableManager,
      options.sourceNode,
      origin
    );
    const values: SerializedCellInRangeResultSnapshot[] = [];

    for (let y = 0; y < relativeSpillArea.height.value; y++) {
      for (let x = 0; x < relativeSpillArea.width.value; x++) {
        const result = captureEvaluationErrors(options.sourceNode, () =>
          evaluation.evaluate({ x, y }, context)
        );
        const serialized = this.serializeSingleEvaluationResultSnapshot(
          result,
          options.getNodeSnapshotId
        );
        if (!serialized) {
          return undefined;
        }
        values.push({
          relativePos: { x, y },
          result: serialized,
        });
      }
    }

    return {
      type: "spilled-values",
      spill: {
        kind: "materialized",
        relativeSpillArea,
        source: evaluation.source,
        sourceCell: evaluation.sourceCell,
        sourceRange: evaluation.sourceRange,
        values,
      },
    };
  }

  deserializeFunctionEvaluationResultSnapshot(
    evaluation: SerializedFunctionEvaluationResultSnapshot,
    resolveNodeSnapshotId: (nodeId: NodeSnapshotId) => DependencyNode
  ): FunctionEvaluationResult {
    if (evaluation.type !== "spilled-values") {
      return this.deserializeSingleEvaluationResultSnapshot(
        evaluation,
        resolveNodeSnapshotId
      );
    }

    const spillSnapshot = evaluation.spill;
    if (spillSnapshot.kind === "materialized") {
      const values = new Map(
        spillSnapshot.values.map((value) => [
          `${value.relativePos.x},${value.relativePos.y}`,
          this.deserializeSingleEvaluationResultSnapshot(
            value.result,
            resolveNodeSnapshotId
          ),
        ])
      );

      return {
        type: "spilled-values",
        source: spillSnapshot.source,
        sourceCell: spillSnapshot.sourceCell,
        sourceRange: spillSnapshot.sourceRange,
        spillArea: (origin) =>
          getAbsoluteRange(spillSnapshot.relativeSpillArea, {
            colIndex: origin.colIndex,
            rowIndex: origin.rowIndex,
          }),
        evaluate: (spillOffset) =>
          values.get(`${spillOffset.x},${spillOffset.y}`) ?? {
            type: "value",
            result: this.convertScalarValueToCellValue(""),
          },
        evaluateAllCells: () => ({
          type: "values",
          values: spillSnapshot.values.map((value) => ({
            relativePos: value.relativePos,
            result: this.deserializeSingleEvaluationResultSnapshot(
              value.result,
              resolveNodeSnapshotId
            ),
          })),
        }),
      };
    }

    return {
      type: "spilled-values",
      source: spillSnapshot.source,
      sourceCell: spillSnapshot.sourceCell,
      sourceRange: spillSnapshot.sourceRange,
      spillArea: (origin) =>
        getAbsoluteRange(spillSnapshot.relativeSpillArea, {
          colIndex: origin.colIndex,
          rowIndex: origin.rowIndex,
        }),
      evaluate: (spillOffset, context) => {
        const cellAddress: CellAddress = {
          workbookName: spillSnapshot.sourceRange.workbookName,
          sheetName: spillSnapshot.sourceRange.sheetName,
          colIndex: spillSnapshot.sourceRange.range.start.col + spillOffset.x,
          rowIndex: spillSnapshot.sourceRange.range.start.row + spillOffset.y,
        };

        const evalNode = this.dependencyManager.getCellValueOrEmptyCellNode(
          cellAddressToKey(cellAddress)
        );
        context.dependencyNode.addDependency(evalNode);
        return evalNode.evaluationResult;
      },
      evaluateAllCells: ({ intersection, context, origin }) => {
        let range = spillSnapshot.sourceRange.range;
        if (intersection) {
          const relativeRange = getRelativeRange(intersection, origin);
          const projectedIntersection = getAbsoluteRange(relativeRange, {
            colIndex: spillSnapshot.sourceRange.range.start.col,
            rowIndex: spillSnapshot.sourceRange.range.start.row,
          });
          const nextRange = getRangeIntersection(
            spillSnapshot.sourceRange.range,
            projectedIntersection
          );
          if (nextRange) {
            range = nextRange;
          }
        }

        const rangeAddress = {
          workbookName: spillSnapshot.sourceRange.workbookName,
          sheetName: spillSnapshot.sourceRange.sheetName,
          range,
        };

        const rangeNode = this.dependencyManager.getRangeNode(
          rangeAddressToKey(rangeAddress)
        );
        context.dependencyNode.addDependency(rangeNode);
        return rangeNode.result;
      },
    };
  }

  serializeSpillMetaEvaluationResultSnapshot(
    evaluation: SpillMetaNode["evaluationResult"],
    options: {
      sourceNode: CellValueNode | SpillMetaNode | AstEvaluationNode;
      getNodeSnapshotId: (node: DependencyNode) => NodeSnapshotId;
    }
  ): SerializedSpillMetaEvaluationResultSnapshot | undefined {
    if (evaluation.type === "does-not-spill") {
      return { type: "does-not-spill" };
    }

    if (evaluation.type === "error") {
      return this.serializeErrorEvaluationResultSnapshot(
        evaluation,
        options.getNodeSnapshotId
      );
    }

    const serialized = this.serializeFunctionEvaluationResultSnapshot(
      evaluation,
      options
    );
    if (!serialized || serialized.type !== "spilled-values") {
      return undefined;
    }

    return serialized;
  }

  deserializeSpillMetaEvaluationResultSnapshot(
    evaluation: SerializedSpillMetaEvaluationResultSnapshot,
    resolveNodeSnapshotId: (nodeId: NodeSnapshotId) => DependencyNode
  ): SpillMetaNode["evaluationResult"] {
    if (evaluation.type === "does-not-spill") {
      return evaluation;
    }

    if (evaluation.type === "error") {
      return this.deserializeSingleEvaluationResultSnapshot(
        evaluation,
        resolveNodeSnapshotId
      ) as SpillMetaNode["evaluationResult"];
    }

    return this.deserializeFunctionEvaluationResultSnapshot(
      evaluation,
      resolveNodeSnapshotId
    ) as SpillMetaNode["evaluationResult"];
  }

  evaluateEmptyCell(node: EmptyCellEvaluationNode): void {
    if (node.resolved) {
      const result = node.evaluationResult;
      if (result && result.type !== "awaiting-evaluation") {
        return;
      }
    }

    this.dependencyManager.unregisterNode(node);
    node.resetDirectDepsUpdated();

    const ctx = new EvaluationContext(
      this.tableManager,
      node,
      node.cellAddress
    );
    const inSpilled = this.dependencyManager.getSpillValue(node.cellAddress);

    if (inSpilled) {
      // if we are spilling then we can just add the spill origin as a dependency and evaluate the spilled value
      const spillTarget = this.dependencyManager.getSpilledAddress(
        node.cellAddress,
        inSpilled
      );
      const spillOriginKey = cellAddressToKey(inSpilled.origin).replace(
        /^[^:]+:/,
        "spill-meta:"
      );
      const spillMetaNode =
        this.dependencyManager.getSpillMetaNode(spillOriginKey);
      node.addDependency(spillMetaNode);
      const result = spillMetaNode.evaluationResult;
      if (result.type === "spilled-values") {
        // let's evaluate the spilled value to extract dependencies
        const evaluation = captureEvaluationErrors(spillMetaNode, () => {
          return result.evaluate(spillTarget.spillOffset, ctx);
        });
        node.setEvaluationResult(evaluation);
      }
    } else {
      // upgrade any frontier dependencies that spill into the range
      node.upgradeFrontierDependencies();

      const evaluationResult: SingleEvaluationResult = {
        type: "value",
        result: this.convertScalarValueToCellValue(""),
      };
      // for now let's just store the empty value, the next time the cell is evaluated isSpilled will be true and the spilled value will be evaluated
      node.setEvaluationResult(evaluationResult);
    }

    this.dependencyManager.registerNode(node);
  }

  evaluateRangeNode(node: RangeEvaluationNode): void {
    if (node.resolved) {
      return;
    }

    this.dependencyManager.unregisterNode(node);
    node.resetDirectDepsUpdated();

    const result = captureEvaluationErrors(node, (): EvaluateAllCellsResult => {
      node.upgradeFrontierDependencies();

      const evalOrder = node.getRangeEvalOrder();
      const circularEntryCache = new Map<DependencyNode, boolean>();
      const shouldSkipCircularEntry = (candidate: DependencyNode) => {
        const cached = circularEntryCache.get(candidate);
        if (cached !== undefined) {
          return cached;
        }

        // If a cell inside the range depends back on the range itself, including
        // its current value would count a circular result back into the same
        // aggregate on a later rerun.
        const isCircular = this.dependsOnNode(candidate, node);
        circularEntryCache.set(candidate, isCircular);
        return isCircular;
      };

      const results: CellInRangeResult[] = [];

      for (const entry of evalOrder) {
        if (entry.type === "value") {
          if (shouldSkipCircularEntry(entry.node)) {
            continue;
          }

          const entryAddress = entry.address;
          const result = entry.node.evaluationResult;

          const relativePos = {
            x: entryAddress.colIndex - node.address.range.start.col,
            y: entryAddress.rowIndex - node.address.range.start.row,
          };

          results.push({ result: result, relativePos });
        } else if (
          entry.type === "empty_cell" ||
          entry.type === "empty_range"
        ) {
          for (const candidateNode of entry.candidates) {
            if (shouldSkipCircularEntry(candidateNode)) {
              continue;
            }

            if (candidateNode.evaluationResult.type === "spilled-values") {
              const spillArea = candidateNode.evaluationResult.spillArea(
                candidateNode.cellAddress
              );
              if (entry.type === "empty_range") {
                const intersects = checkRangeIntersection(
                  spillArea,
                  entry.address.range
                );
                if (intersects) {
                  // When a spilled range intersects with our target range, we need to evaluate
                  // only the cells that fall within the intersection area.
                  //
                  // Example: If cell A10 contains a spilled range that covers A10:B11,
                  // and our target range is B10:INFINITY, then we only want to evaluate
                  // the intersection B10:B11 from the spilled range.
                  //
                  // The evaluateAllCells method expects the intersection to be passed
                  // so it can limit evaluation to only the relevant cells.
                  const ctx = new EvaluationContext(
                    this.tableManager,
                    node,
                    candidateNode.cellAddress
                  );
                  const spilledResults =
                    candidateNode.evaluationResult.evaluateAllCells.call(
                      this.formulaEvaluator,
                      {
                        context: ctx,
                        evaluate: candidateNode.evaluationResult.evaluate,
                        intersection: entry.address.range,
                        origin: candidateNode.cellAddress,
                        lookupOrder: "col-major",
                      }
                    );

                  if (spilledResults.type === "values") {
                    results.push(...spilledResults.values);
                  } else {
                    return spilledResults;
                  }
                }
              } else {
                const intersects = isCellInRange(entry.address, spillArea);
                if (intersects) {
                  // When a spilled range intersects with our target range, we need to evaluate
                  // only the cells that fall within the intersection area.
                  //
                  // Example: If cell A10 contains a spilled range that covers A10:B11,
                  // and our target range is B10:INFINITY, then we only want to evaluate
                  // the intersection B10:B11 from the spilled range.
                  //
                  // The evaluateAllCells method expects the intersection to be passed
                  // so it can limit evaluation to only the relevant cells.

                  const relativePos = {
                    x:
                      entry.address.colIndex -
                      candidateNode.cellAddress.colIndex,
                    y:
                      entry.address.rowIndex -
                      candidateNode.cellAddress.rowIndex,
                  };
                  const ctx = new EvaluationContext(
                    this.tableManager,
                    node,
                    candidateNode.cellAddress
                  );
                  const spilledResult = candidateNode.evaluationResult.evaluate(
                    relativePos,
                    ctx
                  );

                  results.push({
                    relativePos: {
                      x: entry.address.colIndex - node.address.range.start.col,
                      y: entry.address.rowIndex - node.address.range.start.row,
                    },
                    result: spilledResult,
                  });
                }
              }
            }
          }
        }
      }

      return {
        type: "values",
        values: results,
      };
    });

    node.setResult(result);
    this.dependencyManager.registerNode(node);
  }

  evaluateCellNode(
    node: CellValueNode | SpillMetaNode | VirtualCellValueNode
  ): void {
    // Enable caching for resolved nodes
    if (node.resolved) {
      return;
    }

    if (!(node instanceof VirtualCellValueNode)) {
      this.dependencyManager.unregisterNode(node);
    }
    node.resetDirectDepsUpdated();

    const ctx = new EvaluationContext(
      this.tableManager,
      node,
      node.cellAddress
    );

    if (node instanceof CellValueNode && node.spillMeta) {
      // we are evaluating a e.g. A1 in A1=SEQUENCE(10), where we want the value in the cell in A1, i.e. 1
      // As A1 is already pointing to a spill meta node, everything has been setup already,
      // we just need to evaluate the spill origin and assign the result to the currentDepNode
      const spillOrigin = node.spillMeta;
      if (spillOrigin.evaluationResult.type === "spilled-values") {
        const result = spillOrigin.evaluationResult.evaluate(
          { x: 0, y: 0 },
          ctx
        );
        node.setEvaluationResult(result);
        this.dependencyManager.registerNode(node);
        return;
      }
    }

    let content: SerializedCellValue;
    try {
      if (node instanceof VirtualCellValueNode) {
        content = node.cellValue;
      } else {
        content = this.workbookManager.getSerializedCellValue(node.cellAddress);
      }
    } catch (err) {
      const evaluationResult: ErrorEvaluationResult = {
        type: "error",
        err: FormulaError.ERROR,
        message: "Syntax error",
        errAddress: node,
      };
      node.setEvaluationResult(evaluationResult);
      if (!(node instanceof VirtualCellValueNode)) {
        this.dependencyManager.registerNode(node);
      }
      return;
    }

    if (typeof content !== "string" || !content.startsWith("=")) {
      if (node instanceof SpillMetaNode) {
        node.setEvaluationResult({
          type: "does-not-spill",
        });
        this.dependencyManager.registerNode(node);
        return;
      }
      // Static value cells cannot have frontier dependencies
      const result: ValueEvaluationResult = {
        type: "value",
        result: this.convertScalarValueToCellValue(content),
      };
      node.setEvaluationResult(result);
      if (!(node instanceof VirtualCellValueNode)) {
        this.dependencyManager.registerNode(node);
      }
      return;
    }

    let evaluation: FunctionEvaluationResult = captureEvaluationErrors(
      node,
      () => this.formulaEvaluator.evaluateFormula(content.slice(1), ctx)
    );

    // the evaluated cell IS A spilling formula, e.g. if dependencyKey points to A1, then the formula is e.g. A1=SEQUENCE(10), or A1=A3:B5
    if (evaluation.type === "spilled-values") {
      const spillArea = evaluation.spillArea(node.cellAddress);

      if (!this.canSpill(node.cellAddress, spillArea)) {
        // Override evaluation with SPILL error, but continue execution to set up nodes
        evaluation = {
          type: "error",
          err: FormulaError.SPILL,
          message: "Cannot spill - area is blocked",
          errAddress: node,
        };
      } else {
        // Spill succeeds - register it
        this.dependencyManager.setSpilledValue(node.key, {
          spillOnto: spillArea,
          origin: node.cellAddress,
        });
      }

      // Set up spill meta node and evaluation results (even if spill failed)
      if (node instanceof SpillMetaNode) {
        // we have already setup an origin/spill meta node relationship,
        // so we are just reevaluating the spill meta node here
        node.setEvaluationResult(evaluation);
        this.dependencyManager.registerNode(node);
      } else {
        const spillMetaNode = this.dependencyManager.getSpillMetaNode(
          node.key.replace(/^[^:]+:/, "spill-meta:")
        );

        node.addDependency(spillMetaNode);
        node.setSpillMetaNode(spillMetaNode);
        spillMetaNode.setEvaluationResult(evaluation);
        this.dependencyManager.registerNode(spillMetaNode);

        if (evaluation.type === "spilled-values") {
          const originResult = evaluation.evaluate({ x: 0, y: 0 }, ctx);
          node.setEvaluationResult(originResult);
        } else {
          // Spill failed - set error on origin cell
          node.setEvaluationResult(evaluation);
        }
      }
      if (!(node instanceof VirtualCellValueNode)) {
        this.dependencyManager.registerNode(node);
      }
      return;
    }

    if (evaluation.type === "value") {
      if (node instanceof SpillMetaNode) {
        node.setEvaluationResult({
          type: "does-not-spill",
        });
        this.dependencyManager.registerNode(node);
        return;
      } else {
        node.setEvaluationResult(evaluation);
        if (!(node instanceof VirtualCellValueNode)) {
          this.dependencyManager.registerNode(node);
        }
        return;
      }
    }

    node.setEvaluationResult(evaluation);
    if (!(node instanceof VirtualCellValueNode)) {
      this.dependencyManager.registerNode(node);
    }
  }

  evaluateDependencyNode(dependency: DependencyNode): void {
    if (dependency instanceof EmptyCellEvaluationNode) {
      this.evaluateEmptyCell(dependency);
      return;
    }
    if (dependency instanceof RangeEvaluationNode) {
      this.evaluateRangeNode(dependency);
      return;
    }
    if (
      dependency instanceof CellValueNode ||
      dependency instanceof VirtualCellValueNode
    ) {
      this.evaluateCellNode(dependency);
      return;
    }
    if (dependency instanceof AstEvaluationNode) {
      return;
    }
    if (dependency instanceof ResourceDependencyNode) {
      return;
    }
    if (dependency instanceof SpillMetaNode) {
      this.evaluateCellNode(dependency);
      return;
    }
    throw new Error("Invalid dependency: " + (dependency as any).key);
  }

  /**
   * User exposed method to evaluate a formula
   */
  evaluateFormula(
    /**
     * formula for example
     */
    cellValue: SerializedCellValue,
    cellAddress: CellAddress
  ): SerializedCellValue {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }

    const node = this.dependencyManager.getVirtualCellValueNode(
      cellAddress,
      cellValue
    );

    if (node.evaluationResult.type === "awaiting-evaluation") {
      this.evaluateCell(node);
    }

    const result = node.evaluationResult;

    return this.evaluationResultToSerializedValue(result, cellAddress);
  }

  /**
   * Evaluates a cell by building the evaluation order and evaluating the dependencies in order
   */
  evaluateCell(
    node: CellValueNode | EmptyCellEvaluationNode | VirtualCellValueNode
  ): void {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }
    this.isEvaluating = true;

    let precalculatedPlan: EvaluationOrder | undefined;

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;

      // Use DependencyManager to build evaluation order
      const evaluationPlan =
        precalculatedPlan ?? this.dependencyManager.buildEvaluationOrder(node);

      if (evaluationPlan.hasCycle) {
        const evaluationResult: ErrorEvaluationResult = {
          type: "error",
          err: FormulaError.CYCLE,
          message: Array.from(evaluationPlan.cycleNodes ?? [])
            .map((node) => node.key)
            .join(" -> "),
          errAddress: node,
        };
        // cycle detected
        if (evaluationPlan.cycleNodes) {
          for (const node of evaluationPlan.cycleNodes) {
            if (
              !(node instanceof RangeEvaluationNode) &&
              !(node instanceof ResourceDependencyNode)
            ) {
              node.setEvaluationResult(evaluationResult);
            }
          }
        }
        this.isEvaluating = false;
      }

      // Evaluate all dependencies in order
      const timeStart = performance.now();
      const durations: { duration: number; key: string }[] = [];
      let numResolved = 0;
      if (flags.isProfiling && evaluationPlan.evaluationOrder.size > -1) {
        // console.profile();
      }
      evaluationPlan.evaluationOrder.forEach((dependency) => {
        const start = performance.now();
        if (dependency.resolved) {
          numResolved++;
        }
        this.evaluateDependencyNode(dependency);

        const end = performance.now();
        if (flags.isProfiling && evaluationPlan.evaluationOrder.size > -1) {
          durations.push({ duration: end - start, key: dependency.key });
        }
      });
      if (flags.isProfiling && evaluationPlan.evaluationOrder.size > -1) {
        // console.profileEnd();
      }
      if (flags.isProfiling && evaluationPlan.evaluationOrder.size > -1) {
        const percentResolved = Math.round(
          (100 * numResolved) / evaluationPlan.evaluationOrder.size
        );
        const avgDuration =
          durations.reduce((a, b) => a + b.duration, 0) / durations.length || 0;
        console.log(
          `%c[Evaluation] %c${evaluationPlan.evaluationOrder.size} deps | %c${(
            performance.now() - timeStart
          ).toFixed(
            1
          )}ms | %c${percentResolved}% resolved | %c${avgDuration.toFixed(
            2
          )}ms avg`,
          "color:#83aaff;font-weight:bold;",
          "color:#fff;font-weight:bold;",
          "color:#7fff9e",
          "color:#85baff",
          "color:#ffdfa3"
        );
      }

      this.dependencyManager.markResolvedNodes(node);

      const nextEvaluationPlan =
        this.dependencyManager.buildEvaluationOrder(node);

      this.dependencyManager.updateResolvedSCCs(nextEvaluationPlan);

      precalculatedPlan = nextEvaluationPlan;

      // Check if new dependencies were discovered during evaluation
      if (nextEvaluationPlan.hash !== evaluationPlan.hash) {
        requiresReRun = true;
      } else {
        this.isEvaluating = false;
        return;
      }
    }
    this.isEvaluating = false;
  }

  convertScalarValueToCellValue(val: SerializedCellValue): CellValue {
    if (val === "INFINITY") {
      return { type: "infinity", sign: "positive" };
    }
    if (val === "-INFINITY") {
      return { type: "infinity", sign: "negative" };
    }
    if (typeof val === "number") {
      return { type: "number", value: val };
    }
    if (typeof val === "boolean") {
      return { type: "boolean", value: val };
    }
    if (typeof val === "undefined") {
      return { type: "string", value: "" };
    }
    return { type: "string", value: val };
  }

  // todo optimize using workbook manager
  canSpill(spillCandidate: CellAddress, spillArea: SpreadsheetRange): boolean {
    // Check if the spill origin cell is inside a table - spilling formulas cannot exist in tables
    if (this.tableManager.isCellInTable(spillCandidate)) {
      return false;
    }

    // Check if the spill area would intersect with any table - formulas cannot spill into tables
    if (
      this.tableManager.doesRangeIntersectTable(
        spillCandidate.workbookName,
        spillCandidate.sheetName,
        spillArea
      )
    ) {
      return false;
    }

    const sheet = this.workbookManager.getSheet(spillCandidate);
    if (!sheet) {
      throw new SheetNotFoundError(spillCandidate.sheetName);
    }
    const cellId = getCellReference(spillCandidate);
    const content = sheet.content.get(cellId);
    if (!content) {
      throw new EvaluationError(FormulaError.REF, `Cell not found: ${cellId}`);
    }
    for (const spilledValue of this.dependencyManager.spilledValues) {
      if (
        spilledValue.origin.workbookName !== spillCandidate.workbookName ||
        spilledValue.origin.sheetName !== spillCandidate.sheetName
      ) {
        continue;
      }
      if (
        spilledValue.origin.colIndex === spillCandidate.colIndex &&
        spilledValue.origin.rowIndex === spillCandidate.rowIndex
      ) {
        // we are already have a spill, this one will be replaced
        continue;
      }

      if (checkRangeIntersection(spillArea, spilledValue.spillOnto)) {
        return false;
      }
    }
    // let's just check the raw data if there is something in the range
    for (const key of sheet.content.keys()) {
      const cellAddress = parseCellReference(key);
      const endCol = spillArea.end.col;
      const endRow = spillArea.end.row;

      if (
        cellAddress.colIndex === spillCandidate.colIndex &&
        cellAddress.rowIndex === spillCandidate.rowIndex
      ) {
        continue;
      }

      if (endCol.type === "number" && endRow.type === "number") {
        if (
          cellAddress.colIndex >= spillArea.start.col &&
          cellAddress.colIndex <= endCol.value &&
          cellAddress.rowIndex >= spillArea.start.row &&
          cellAddress.rowIndex <= endRow.value
        ) {
          if (
            normalizeSerializedCellValue(sheet.content.get(key)) !== undefined
          ) {
            // there is something in the range, so we can't spill
            return false;
          }
        }
      }
    }

    return true;
  }

  getCellEvaluationResult(
    cellAddress: CellAddress
  ): SingleEvaluationResult | undefined {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }

    const nodeKey = cellAddressToKey(cellAddress);
    const node = this.dependencyManager.getCellValueOrEmptyCellNode(nodeKey);

    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      throw new SheetNotFoundError(cellAddress.sheetName);
    }

    if (node.evaluationResult.type === "awaiting-evaluation") {
      // if (cellAddressToKey(cellAddress).includes("G10")) {
      //   console.group("Evaluation of G10");
      //   flags.isProfiling = true;
      //   console.time("Evaluation of G10");
      //   console.profile("Evaluation of G10");
      // }
      this.evaluateCell(node);
      // if (flags.isProfiling) {
      //   flags.isProfiling = false;
      //   console.timeEnd("Evaluation of G10");
      //   console.profileEnd("Evaluation of G10");
      //   console.groupEnd();
      // }
    }

    const result = node.evaluationResult;

    return result;
  }
}
