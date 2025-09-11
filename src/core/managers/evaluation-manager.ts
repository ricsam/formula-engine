import { normalizeSerializedCellValue } from "src/parser/formatter";
import { FormulaEvaluator } from "../../evaluator/formula-evaluator";
import {
  FormulaError,
  type CellAddress,
  type CellValue,
  type EvaluatedDependencyNode,
  type EvaluationContext,
  type FunctionEvaluationResult,
  type NamedExpression,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpilledValue,
  type SpreadsheetRange,
  type TableDefinition,
} from "../types";
import {
  getCellReference,
  isCellInRange,
  isRangeOneCell,
  parseCellReference,
} from "../utils";
import {
  dependencyNodeToKey,
  keyToDependencyNode,
} from "../utils/dependency-node-key";
import type { NamedExpressionManager } from "./named-expression-manager";
import type { WorkbookManager } from "./workbook-manager";

export class EvaluationManager {
  private isEvaluating = false;

  evaluatedNodes: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey
     */
    string,
    EvaluatedDependencyNode
  > = new Map();

  spilledValues: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey for the origin cell
     */
    string,
    SpilledValue
  > = new Map();

  constructor(
    private workbookManager: WorkbookManager,
    private namedExpressionManager: NamedExpressionManager,
    private formulaEvaluator: FormulaEvaluator
  ) {}

  getEvaluatedNodes() {
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
    evaluation: SingleEvaluationResult,
    debug?: boolean
  ): SerializedCellValue {
    if (evaluation.type !== "error") {
      const value = evaluation.result;

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

  getNodeDeps(nodeKey: string): Set<string> {
    const deps = new Set<string>();
    const node = this.evaluatedNodes.get(nodeKey);
    node?.deps?.forEach((dep) => deps.add(dep));
    node?.frontierDependencies?.forEach((frontierDep) => {
      if (node?.discardedFrontierDependencies?.has(frontierDep)) {
        return;
      }
      return deps.add(frontierDep);
    });
    return deps;
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

      const deps = this.getNodeDeps(current);

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
      const deps = this.getNodeDeps(node);
      for (const precedent of deps) {
        inDegree.set(precedent, (inDegree.get(precedent) || 0) + 1);
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

      const deps = this.getNodeDeps(current);

      for (const dependent of deps) {
        const degree = inDegree.get(dependent)! - 1;
        inDegree.set(dependent, degree);

        if (degree === 0) {
          queue.push(dependent);
        }
      }
    }

    // Check if all nodes were processed (no cycles)
    if (result.length !== nodeKeys.size) {
      return null; // Cycle detected
    }

    return result;
  }

  isSpilled(cellAddress: CellAddress): SpilledValue | undefined {
    for (const spilledValue of this.spilledValues.values()) {
      if (spilledValue.origin.sheetName !== cellAddress.sheetName) {
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

  evaluateSpilled(
    cellAddress: CellAddress,
    context: EvaluationContext
  ):
    | { isSpilled: true; result: FunctionEvaluationResult | undefined }
    | { isSpilled: false } {
    const spilled = this.isSpilled(cellAddress);
    if (spilled) {
      const spillSource = this.getSpilledAddress(cellAddress, spilled);
      const spillOrigin = this.evalTimeSafeEvaluateCell(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return {
          isSpilled: true,
          result: spillOrigin.evaluate(spillSource.spillOffset, context),
        };
      }
    }
    return { isSpilled: false };
  }

  /**
   * During evaluation, we can't use the formula evaluator to evaluate a cell, because it will create a cycle.
   * This method can be used to "evaluate" a cell during evaluation, without creating a cycle.
   *
   * Internally this method will try to look up the evaluated result for the cell,
   * if it doesn't exist it will push the cell address to the dependency graph,
   * causing the engine to re-evaluate the dependency graph,
   * such that on the second evaluation the cell's evaluation result will be available.
   */
  evalTimeSafeEvaluateCell(
    cellAddress: CellAddress,
    context: EvaluationContext
  ): FunctionEvaluationResult | undefined {
    const spilled = this.isSpilled(cellAddress);
    if (spilled) {
      const spillSource = this.getSpilledAddress(cellAddress, spilled);
      const spillOrigin = this.evalTimeSafeEvaluateCell(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return spillOrigin.evaluate(spillSource.spillOffset, context);
      }
    }
    const key = dependencyNodeToKey({
      type: "cell",
      address: cellAddress,
      sheetName: cellAddress.sheetName,
      workbookName: cellAddress.workbookName,
    });
    context.dependencies.add(key);
    return this.evaluatedNodes.get(key)?.evaluationResult;
  }

  /**
   * Similar logic as evalTimeSafeEvaluateCell, but for named expressions
   */
  evalTimeSafeEvaluateNamedExpression(
    namedExpression: Pick<
      NamedExpression,
      "name" | "sheetName" | "workbookName"
    >,
    context: EvaluationContext
  ): FunctionEvaluationResult | undefined {
    const nodeKey = dependencyNodeToKey({
      type: "named-expression",
      name: namedExpression.name,
      sheetName: namedExpression.sheetName ?? context.currentSheet,
      workbookName: namedExpression.workbookName ?? context.currentWorkbook,
    });
    context.dependencies.add(nodeKey);

    const value = this.evaluatedNodes.get(nodeKey);
    return value?.evaluationResult;
  }

  getSpilledAddress(
    cellAddress: CellAddress,
    /**
     * if the spilled value is already available, we can use it to get the source address
     */
    passedSpilledValue?: SpilledValue
  ): { address: CellAddress; spillOffset: { x: number; y: number } } {
    const spilledValue = passedSpilledValue ?? this.isSpilled(cellAddress);
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
    const node = keyToDependencyNode(nodeKey);

    const dependenciesDiscoveredInEvaluation: Set<string> = new Set();
    const frontierDependenciesDiscoveredInEvaluation: Set<string> = new Set();
    const discardedFrontierDependenciesDiscoveredInEvaluation: Set<string> =
      new Set();

    let evaluation: FunctionEvaluationResult | undefined;

    if (node.type === "cell") {
      const cellId = getCellReference({
        rowIndex: node.address.rowIndex,
        colIndex: node.address.colIndex,
      });

      const nodeAddress: CellAddress = {
        workbookName: node.workbookName,
        sheetName: node.sheetName,
        colIndex: node.address.colIndex,
        rowIndex: node.address.rowIndex,
      };

      const sheet = this.workbookManager.getSheet(cellAddress);

      if (!sheet) {
        this.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "error",
            err: FormulaError.REF,
            message: "Sheet not found",
          },
        });
        return false;
      }

      let content: SerializedCellValue;
      try {
        content = normalizeSerializedCellValue(sheet.content.get(cellId));
      } catch (err) {
        this.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "error",
            err: FormulaError.ERROR,
            message: "Syntax error",
          },
        });
        return false;
      }

      const evaluationContext: EvaluationContext = {
        currentSheet: sheet.name,
        currentWorkbook: node.workbookName,
        currentCell: nodeAddress,
        evaluationStack: new Set(),
        dependencies: dependenciesDiscoveredInEvaluation,
        frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
        discardedFrontierDependencies:
          discardedFrontierDependenciesDiscoveredInEvaluation,
      };

      if (typeof content !== "string" || !content.startsWith("=")) {
        const spilled = this.isSpilled(nodeAddress);
        if (spilled) {
          const spillTarget = this.getSpilledAddress(nodeAddress, spilled);
          const spillOrigin = this.evalTimeSafeEvaluateCell(
            spilled.origin,
            evaluationContext
          );
          if (spillOrigin && spillOrigin.type === "spilled-values") {
            // let's evaluate the spilled value to extract dependencies
            evaluation = spillOrigin.evaluate(
              spillTarget.spillOffset,
              evaluationContext
            );
          }
        } else {
          this.evaluatedNodes.set(nodeKey, {
            evaluationResult: {
              type: "value",
              result: this.convertScalarValueToCellValue(content),
            },
          });
          return false;
        }
      } else {
        evaluation = this.formulaEvaluator.evaluateFormula(
          content.slice(1),
          evaluationContext
        );
      }

      // if a cell returns a range, we need to spill the values onto the sheet
      if (evaluation && evaluation.type === "spilled-values") {
        const spillArea = evaluation.spillArea(nodeAddress);
        if (!isRangeOneCell(spillArea)) {
          if (this.canSpill(nodeAddress, spillArea)) {
            this.spilledValues.set(nodeKey, {
              spillOnto: spillArea,
              origin: nodeAddress,
            });

            this.evaluatedNodes.forEach((evaled, key) => {
              const isDependencyInRange = (dep: string) => {
                const node = keyToDependencyNode(dep);
                if (node.type === "cell") {
                  const cellAddress: CellAddress = {
                    ...node.address,
                    sheetName: node.sheetName,
                    workbookName: node.workbookName,
                  };
                  return isCellInRange(cellAddress, spillArea);
                }
                return false;
              };

              const frontierDeps = new Set(
                evaled.frontierDependencies ?? []
              ).union(new Set(evaled.discardedFrontierDependencies ?? []));

              for (const dep of frontierDeps) {
                if (isDependencyInRange(dep)) {
                  // one of the transient frontier dependencies of key is in the spill area,
                  // we need to re-evaluate the cell, and potentially previously discarded frontier dependencies
                  // could now be frontier dependencies
                  // e.g. if a fronteir dependency was dependant on a spilled cell,
                  // previously the spilled cell was just "" but after spilling it
                  // gets a value making the frontier dependency potencially spill onto the spill area
                  // making it a frontier dependency
                  evaled.discardedFrontierDependencies = new Set();
                  this.evaluateDependencyNode(key, cellAddress);
                  return; // go to next evaluated node
                }
              }

              for (const dep of evaled.deps ?? []) {
                if (isDependencyInRange(dep)) {
                  // one of the dependencies of key is in the spill area,
                  // we need to re-evaluate the cell
                  this.evaluateDependencyNode(key, cellAddress);
                  return; // go to next evaluated node
                }
              }
            });
          } else {
            evaluation = {
              type: "error",
              err: FormulaError.SPILL,
              message: "Can't spill",
            };
          }
        }
      }
    } else if (node.type === "named-expression") {
      const expression = this.namedExpressionManager.getNamedExpression({
        sheetName: node.sheetName,
        workbookName: node.workbookName,
        name: node.name,
      });
      if (!expression) {
        this.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "error",
            err: FormulaError.NAME,
            message: "Named expression not found",
          },
        });
        return false;
      }

      evaluation = this.formulaEvaluator.evaluateFormula(
        expression.expression,
        {
          currentSheet: cellAddress.sheetName,
          currentWorkbook: cellAddress.workbookName,
          currentCell: cellAddress,
          evaluationStack: new Set(),
          dependencies: dependenciesDiscoveredInEvaluation,
          frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
          discardedFrontierDependencies:
            discardedFrontierDependenciesDiscoveredInEvaluation,
        }
      );
    } else {
      throw new Error(`${node.type} is not supported yet in the evaluator`);
    }

    if (!evaluation) {
      throw new Error(`${node.type} is not supported yet in the evaluator`);
    }

    const currentDeps = this.evaluatedNodes.get(nodeKey)?.deps ?? new Set();
    const currentFrontierDeps =
      this.evaluatedNodes.get(nodeKey)?.frontierDependencies ?? new Set();

    let requiresReRun = true;
    if (
      !(
        dependenciesDiscoveredInEvaluation.isSubsetOf(currentDeps) &&
        currentDeps.isSubsetOf(dependenciesDiscoveredInEvaluation)
      ) ||
      currentDeps.size !== dependenciesDiscoveredInEvaluation.size
    ) {
      requiresReRun = true;
    }
    if (
      !(
        frontierDependenciesDiscoveredInEvaluation.isSubsetOf(
          currentFrontierDeps
        ) &&
        currentFrontierDeps.isSubsetOf(
          frontierDependenciesDiscoveredInEvaluation
        )
      ) ||
      currentFrontierDeps.size !==
        frontierDependenciesDiscoveredInEvaluation.size
    ) {
      requiresReRun = true;
    }

    this.evaluatedNodes.set(nodeKey, {
      deps: dependenciesDiscoveredInEvaluation,
      frontierDependencies: frontierDependenciesDiscoveredInEvaluation,
      discardedFrontierDependencies:
        discardedFrontierDependenciesDiscoveredInEvaluation,
      evaluationResult: evaluation,
    });

    return requiresReRun;
  }

  evaluateCell(cellAddress: CellAddress): void {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }
    this.isEvaluating = true;
    const sheet = this.workbookManager.getSheet(cellAddress);
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
      sheetName: cellAddress.sheetName,
      workbookName: cellAddress.workbookName,
    });

    let requiresReRun = true;
    while (requiresReRun) {
      requiresReRun = false;
      let content: SerializedCellValue;
      try {
        content = normalizeSerializedCellValue(sheet.content.get(cellId));
      } catch (err) {
        this.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "error",
            err: FormulaError.ERROR,
            message: "Syntax error",
          },
        });
        break;
      }
      if (typeof content !== "string" || !content.startsWith("=")) {
        this.evaluatedNodes.set(nodeKey, {
          evaluationResult: {
            type: "value",
            result: this.convertScalarValueToCellValue(content),
          },
        });
        break;
      }

      const allDeps = this.getTransitiveDeps(nodeKey);
      const sorted = this.topologicalSort(allDeps)?.reverse();

      if (!sorted) {
        // cycle detected
        this.evaluatedNodes.set(nodeKey, {
          deps: this.evaluatedNodes.get(nodeKey)?.deps ?? new Set(),
          frontierDependencies:
            this.evaluatedNodes.get(nodeKey)?.frontierDependencies ?? new Set(),
          discardedFrontierDependencies:
            this.evaluatedNodes.get(nodeKey)?.discardedFrontierDependencies ??
            new Set(),
          evaluationResult: {
            type: "error",
            err: FormulaError.CYCLE,
            message: "Cycle detected",
          },
        });
        this.isEvaluating = false;
        return;
      }

      sorted.forEach((nodeKey) =>
        this.evaluateDependencyNode(nodeKey, cellAddress)
      );
      this.evaluateDependencyNode(nodeKey, cellAddress);

      const transitiveDeps2 = this.getTransitiveDeps(nodeKey);

      // the cells were potentially evaluated in the wrong order
      if (
        allDeps.size !== transitiveDeps2.size ||
        !allDeps.isSubsetOf(transitiveDeps2)
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
    const sheet = this.workbookManager.getSheet(originCellAddress);
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
      if (isCellInRange(originCellAddress, value.spillOnto)) {
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
  ): SingleEvaluationResult | undefined {
    if (this.isEvaluating) {
      throw new Error("Evaluation in progress");
    }

    const sheet = this.workbookManager.getSheet(cellAddress);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // maybe it is a spilled cell, we need to check the spilled values
    // the context is quite irrelevant, because the cells are "cached" if spilled is true
    const dummyContext: EvaluationContext = {
      currentSheet: cellAddress.sheetName,
      currentWorkbook: cellAddress.workbookName,
      currentCell: cellAddress,
      evaluationStack: new Set(),
      dependencies: new Set(),
      frontierDependencies: new Set(),
      discardedFrontierDependencies: new Set(),
    };
    const spilled = this.evaluateSpilled(cellAddress, dummyContext);
    if (spilled.isSpilled) {
      const result = spilled.result;
      if (!result) {
        return undefined;
      }
      if (result.type === "spilled-values") {
        return result.evaluate({ x: 0, y: 0 }, dummyContext);
      }
      return result;
    }

    this.evaluateCell(cellAddress);

    const value = this.evaluatedNodes.get(
      dependencyNodeToKey({
        type: "cell",
        address: cellAddress,
        sheetName: cellAddress.sheetName,
        workbookName: cellAddress.workbookName,
      })
    );

    if (!value || !value.evaluationResult) {
      // nothing in the cell
      return undefined;
    }

    const result = value.evaluationResult;
    if (result.type === "spilled-values") {
      return result.evaluate({ x: 0, y: 0 }, dummyContext);
    }
    return result;
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.formulaEvaluator.isCellInTable(cellAddress);
  }
}
