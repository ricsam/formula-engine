import { functions } from "src/functions";
import type { FunctionNode } from "src/parser/ast";
import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type EvaluationContext,
  type FunctionEvaluationResult,
  type NamedExpression,
  type SerializedCellValue,
  type SpilledValue,
  type SpreadsheetRange,
  type TableDefinition,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";
import {
  dependencyNodeToKey,
  keyToDependencyNode,
} from "../utils/dependency-node-key";

export class EvaluationManager extends FormulaEvaluator {
  private isEvaluating = false;

  constructor(
    sheets: Map<string, any>,
    scopedNamedExpressions: Map<string, Map<string, NamedExpression>>,
    globalNamedExpressions: Map<string, NamedExpression>,
    tables: Map<string, TableDefinition>
  ) {
    super();
    this.sheets = sheets;
    this.scopedNamedExpressions = scopedNamedExpressions;
    this.globalNamedExpressions = globalNamedExpressions;
    this.tables = tables;
  }

  getEvaluatedNodes(): Map<
    string,
    { deps: Set<string>; evaluationResult?: FunctionEvaluationResult }
  > {
    return this.evaluatedNodes;
  }

  getSpilledValues(): Map<string, SpilledValue> {
    return this.spilledValues;
  }

  clearEvaluationCache(): void {
    this.evaluatedNodes.clear();
    this.spilledValues.clear();
  }

  evaluationResultToSerializedValue(
    evaluation: FunctionEvaluationResult,
    debug?: boolean
  ): SerializedCellValue {
    if (evaluation.type !== "error") {
      const value =
        evaluation.type === "value"
          ? evaluation.result
          : evaluation.originResult;

      return value.type === "infinity"
        ? value.sign === "positive"
          ? "INFINITY"
          : "-INFINITY"
        : value.value;
    }

    if (debug) {
      return evaluation.err + ": " + evaluation.message;
    }

    return evaluation.err;
  }

  getTransitiveDeps(
    /**
     * nodeKey is the dependency node key, from dependencyNodeToKey
     */
    nodeKey: string
  ): Set<string> {
    const visited = new Set<string>();
    const queue = [nodeKey];

    while (queue.length > 0) {
      const current = queue.shift()!;

      if (visited.has(current)) continue;
      visited.add(current);

      const deps = this.evaluatedNodes.get(current)?.deps;

      if (!deps) {
        continue;
      }

      for (const dep of deps) {
        queue.push(dep);
      }
    }

    visited.delete(nodeKey); // Don't include the starting node
    return visited;
  }

  private topologicalSort(
    /**
     * nodeKeys is the set of dependency node keys, see dependencyNodeToKey
     */
    nodeKeys: Set<string>
  ): string[] | null {
    const inDegree = new Map<string, number>();
    const queue: string[] = [];
    const result: string[] = [];

    // Calculate in-degrees
    for (const node of nodeKeys) {
      inDegree.set(node, 0);
    }

    for (const node of nodeKeys) {
      const deps = this.evaluatedNodes.get(node)?.deps;
      if (deps) {
        for (const precedent of deps) {
          inDegree.set(precedent, (inDegree.get(precedent) || 0) + 1);
        }
      }
    }

    // Find nodes with no incoming edges
    for (const [node, degree] of inDegree) {
      if (degree === 0) {
        queue.push(node);
      }
    }

    // Process queue
    while (queue.length > 0) {
      const current = queue.shift()!;
      result.push(current);

      const deps = this.evaluatedNodes.get(current)?.deps;

      if (deps) {
        for (const dependent of deps) {
          const degree = inDegree.get(dependent)! - 1;
          inDegree.set(dependent, degree);

          if (degree === 0) {
            queue.push(dependent);
          }
        }
      }
    }

    // Check if all nodes were processed (no cycles)
    if (result.length !== nodeKeys.size) {
      return null; // Cycle detected
    }

    return result;
  }

  evaluateDependencyNode(
    /**
     * nodeKey is the dependency node key, from dependencyNodeToKey
     */
    nodeKey: string,
    /**
     * We evaluate the dependency node in the context of the cell address
     */
    cellAddress: CellAddress
  ): boolean {
    let requiresReRun = true;

    const node = keyToDependencyNode(nodeKey);

    const dependenciesDiscoveredInEvaluation: Set<string> = new Set();

    let evaluation: FunctionEvaluationResult | undefined;

    if (node.type === "cell") {
      const cellId = getCellReference({
        rowIndex: node.address.rowIndex,
        colIndex: node.address.colIndex,
      });

      const nodeAddress: CellAddress = {
        sheetName: node.sheetName,
        colIndex: node.address.colIndex,
        rowIndex: node.address.rowIndex,
      };

      const sheet = this.sheets.get(node.sheetName);

      if (!sheet) {
        this.evaluatedNodes.set(nodeKey, {
          deps: new Set(),
          evaluationResult: {
            type: "error",
            err: FormulaError.REF,
            message: "Sheet not found",
          },
        });
        return requiresReRun;
      }

      const content = normalizeSerializedCellValue(sheet.content.get(cellId));

      const evaluationContext: EvaluationContext = {
        currentSheet: sheet.name,
        currentCell: nodeAddress,
        evaluationStack: new Set(),
        dependencies: dependenciesDiscoveredInEvaluation,
      };

      if (typeof content !== "string" || !content.startsWith("=")) {
        const spilled = this.isSpilled(nodeAddress);
        if (spilled) {
          const spillTarget = this.getSpilledAddress(nodeAddress, spilled);
          const spillOrigin = this.runtimeSafeEvaluatedNode(
            spilled.origin,
            evaluationContext
          );
          if (spillOrigin && spillOrigin.type === "spilled-values") {
            // let's evaluate the spilled value to extract dependencies
            evaluation = spillOrigin.evaluate(spillTarget, evaluationContext);
          }
        } else {
          this.evaluatedNodes.set(nodeKey, {
            deps: new Set(),
            evaluationResult: {
              type: "value",
              result: this.convertScalarValueToCellValue(content),
            },
          });
          return requiresReRun;
        }
      } else {
        evaluation = this.evaluateFormula(content.slice(1), evaluationContext);
      }

      // Optimization: if spilled values are one cell, then change the evaluation to a value

      // if a cell returns a range, we need to spill the values onto the sheet
      if (
        evaluation &&
        evaluation.type === "spilled-values" &&
        !this.isRangeOneCell(evaluation.spillArea)
      ) {
        if (this.canSpill(nodeAddress, evaluation.spillArea)) {
          this.spilledValues.set(nodeKey, {
            spillOnto: evaluation.spillArea,
            origin: nodeAddress,
          });
        } else {
          evaluation = {
            type: "error",
            err: FormulaError.SPILL,
            message: "Can't spill",
          };
        }
      }
    } else if (node.type === "named-expression") {
      const expression =
        this.scopedNamedExpressions.get(node.sheetName)?.get(node.name) ??
        this.globalNamedExpressions.get(node.name);
      if (!expression) {
        this.evaluatedNodes.set(nodeKey, {
          deps: new Set(),
          evaluationResult: {
            type: "error",
            err: FormulaError.NAME,
            message: "Named expression not found",
          },
        });
        return requiresReRun;
      }

      evaluation = this.evaluateFormula(expression.expression, {
        currentSheet: cellAddress.sheetName,
        currentCell: cellAddress,
        evaluationStack: new Set(),
        dependencies: dependenciesDiscoveredInEvaluation,
      });
    } else {
      throw new Error(`${node.type} is not supported yet in the evaluator`);
    }

    if (!evaluation) {
      throw new Error(`${node.type} is not supported yet in the evaluator`);
    }

    const currentDeps = this.evaluatedNodes.get(nodeKey)?.deps ?? new Set();
    if (
      !(
        dependenciesDiscoveredInEvaluation.isSubsetOf(currentDeps) &&
        currentDeps.isSubsetOf(dependenciesDiscoveredInEvaluation)
      ) ||
      currentDeps.size !== dependenciesDiscoveredInEvaluation.size
    ) {
      requiresReRun = true;
    }

    this.evaluatedNodes.set(nodeKey, {
      deps: dependenciesDiscoveredInEvaluation,
      evaluationResult: evaluation,
    });

    return requiresReRun;
  }

  evaluateCell(cellAddress: CellAddress): void {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }
    this.isEvaluating = true;
    const sheet = this.sheets.get(cellAddress.sheetName);
    if (!sheet) {
      this.isEvaluating = false;
      throw new Error("Sheet not found");
    }

    const cellId = getCellReference({
      rowIndex: cellAddress.rowIndex,
      colIndex: cellAddress.colIndex,
    });

    const nodeKey = dependencyNodeToKey({
      type: "cell",
      address: cellAddress,
      sheetName: sheet.name,
    });

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;
      const content = normalizeSerializedCellValue(sheet.content.get(cellId));
      if (typeof content !== "string" || !content.startsWith("=")) {
        this.evaluatedNodes.set(nodeKey, {
          deps: new Set(),
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        });
        break;
      }

      const allDeps = this.getTransitiveDeps(
        dependencyNodeToKey({
          type: "cell",
          address: cellAddress,
          sheetName: sheet.name,
        })
      );

      const sorted = this.topologicalSort(allDeps)?.reverse();
      if (!sorted) {
        // cycle detected
        this.evaluatedNodes.set(nodeKey, {
          deps: allDeps,
          evaluationResult: {
            type: "error",
            err: FormulaError.CYCLE,
            message: "Cycle detected",
          },
        });
        this.isEvaluating = false;
        return;
      }

      const transitiveDeps1 = this.getTransitiveDeps(nodeKey);

      sorted.forEach((nodeKey) =>
        this.evaluateDependencyNode(nodeKey, cellAddress)
      );
      this.evaluateDependencyNode(nodeKey, cellAddress);

      const transitiveDeps2 = this.getTransitiveDeps(nodeKey);

      // the cells were potentially evaluated in the wrong order
      if (
        transitiveDeps1.size !== transitiveDeps2.size ||
        !transitiveDeps1.isSubsetOf(transitiveDeps2)
      ) {
        requiresReRun = true;
      }
    }
    this.isEvaluating = false;
  }

  convertScalarValueToCellValue(
    val: undefined | boolean | number | string
  ): CellValue {
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

  canSpill(originCellAddress: CellAddress, range: SpreadsheetRange): boolean {
    const sheet = this.sheets.get(originCellAddress.sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }
    const cellId = getCellReference(originCellAddress);
    const content = sheet.content.get(cellId);
    if (!content) {
      throw new Error(`Cell not found: ${cellId}`);
    }
    // let's just check the raw data if there is something in the range
    for (const key of sheet.content.keys()) {
      const cellAddress = parseCellReference(key);
      const endCol = range.end.col;
      const endRow = range.end.row;

      if (
        cellAddress.colIndex === originCellAddress.colIndex &&
        cellAddress.rowIndex === originCellAddress.rowIndex
      ) {
        continue;
      }

      if (endCol.type === "number" && endRow.type === "number") {
        if (
          cellAddress.colIndex >= range.start.col &&
          cellAddress.colIndex <= endCol.value &&
          cellAddress.rowIndex >= range.start.row &&
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
    for (const value of this.spilledValues.values()) {
      if (this.isCellInRange(originCellAddress, value.spillOnto)) {
        if (
          value.origin.colIndex === originCellAddress.colIndex &&
          value.origin.rowIndex === originCellAddress.rowIndex
        ) {
          continue;
        }
        return false;
      }
    }

    return true;
  }

  getCellEvaluationResult(
    cellAddress: CellAddress
  ): FunctionEvaluationResult | undefined {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }

    const sheet = this.sheets.get(cellAddress.sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // maybe it is a spilled cell, we need to check the spilled values
    const context: EvaluationContext = {
      currentSheet: cellAddress.sheetName,
      currentCell: cellAddress,
      evaluationStack: new Set(),
      dependencies: new Set(),
    };
    const spilled = this.evaluateSpilled(cellAddress, context);
    if (spilled.isSpilled) {
      return spilled.result;
    }

    this.evaluateCell(cellAddress);

    const value = this.evaluatedNodes.get(
      dependencyNodeToKey({
        type: "cell",
        address: cellAddress,
        sheetName: sheet.name,
      })
    );

    if (!value || !value.evaluationResult) {
      // nothing in the cell
      return undefined;
    }

    return value.evaluationResult;
  }
}
