import type { TableManager } from "src/core/managers";
import type { DependencyNode } from "src/core/managers/dependency-node";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { CellAddress } from "src/core/types";

export class EvaluationContext {
  private _cellAddress: CellAddress;
  private _tableName: string | undefined;
  /**
   * Can be a range or a cell
   */
  private _dependencyNode: DependencyNode;
  constructor(
    tableManager: TableManager,
    dependencyNode: DependencyNode,
    cellAddress: CellAddress
  ) {
    this._dependencyNode = dependencyNode;
    this._cellAddress = cellAddress;
    const table = tableManager.isCellInTable(cellAddress);
    this._tableName = table?.name;
  }

  get dependencyNode() {
    return this._dependencyNode;
  }

  private _contextDependency: ContextDependency = {};

  getContextDependency() {
    return this._contextDependency;
  }

  /**
   * The cell context, the address of the cell being evaluated
   * and the context in which results should be stored
   */
  get cellAddress() {
    return this._cellAddress;
  }

  get tableName() {
    return this._tableName;
  }

  addContextDependency(...types: ContextDependencyType[]) {
    for (const type of types) {
      switch (type) {
        case "row":
          this._contextDependency.rowIndex = this._cellAddress.rowIndex;
          break;
        case "col":
          this._contextDependency.colIndex = this._cellAddress.colIndex;
          break;
        case "workbook":
          this._contextDependency.workbookName = this._cellAddress.workbookName;
          break;
        case "sheet":
          this._contextDependency.sheetName = this._cellAddress.sheetName;
          break;
        case "table":
          this._contextDependency.tableName = this.tableName;
          break;
        default:
          throw new Error(`Invalid context dependency type: ${type}`);
      }
    }
  }

  /**
   * When evaluating an AST node,
   * we need to append the subtree context
   * dependencies to the current context dependency
   */
  appendContextDependency(contextDependency: ContextDependency) {
    this._contextDependency = {
      ...this._contextDependency,
      ...Object.fromEntries(
        Object.entries(contextDependency).filter(
          ([key, value]) => value !== undefined
        )
      ),
    };
  }
}

/**
 * Each value has the same value as the origin cell
 * the defined keys are the ones the ast node is dependent on
 * e.g. A3=ROW() will have a context dependency of { rowIndex: 3 }
 *
 * The keys are ANDed together, e.g. { workbookName: "Sheet1", sheetName: "Sheet2" }
 * means the ast node is dependent on the workbook "Sheet1" and the sheet "Sheet2"
 */
export type ContextDependency = {
  workbookName?: string;
  sheetName?: string;
  tableName?: string;
  rowIndex?: number;
  colIndex?: number;
};

export const contextDependencyKeys = [
  "workbookName",
  "sheetName",
  "tableName",
  "rowIndex",
  "colIndex",
] as const;

/**
 * These are some distinct scenarios where context dependencies are added
 */
export type ContextDependencyType =
  | "row"
  | "col"
  | "workbook"
  | "sheet"
  | "table";

// *  [astKey], // `=1+1`
// *  [astKey, sheetKey, workbookKey], // `B3`
// *  [astKey, workbookKey], // `Table1[Column1]`
// *  [astKey, workbookKey], // `Sheet1!B3`
// *  [astKey, cellAddress.rowIndex], // `ROW()`
// *  [astKey, cellAddress.colIndex], // `COL()`
// *  [astKey, cellAddress.rowIndex, cellAddress.colIndex] // `CELL("address")`
// *  [astKey, tableKey, cellAddress.rowIndex], // `@Column1`
// *  [astKey, workbookKey, cellAddress.rowIndex], // `Table1[@Column1]`

type Dim = "workbookName" | "sheetName" | "tableName" | "rowIndex" | "colIndex";
const DIM_ORDER: readonly Dim[] = [
  "workbookName",
  "sheetName",
  "tableName",
  "rowIndex",
  "colIndex",
] as const;

/** Build the canonical cache key string for a dependency. */
export function keyFromDependency(dep: ContextDependency): string {
  return DIM_ORDER.map((k) =>
    dep[k] !== undefined ? `${k}:[${(dep as any)[k]}]` : `${k}:*`
  ).join(",");
}

type Context = {
  workbookName: string;
  sheetName: string;
  tableName?: string;
  rowIndex: number;
  colIndex: number;
};

/** True if `dep` matches `ctx` (i.e., every specified field equals the context’s field). */
export function dependencyMatchesContext(
  ctx: Context,
  dep: ContextDependency
): boolean {
  return DIM_ORDER.every(
    (k) => dep[k] === undefined || (dep as any)[k] === (ctx as any)[k]
  );
}

const contextCacheKey = new Map<string, string[]>();

/**
 * Generate every cache key that would be eligible for `ctx`.
 * By default returns keys ordered from most-specific to least-specific.
 */
export function eligibleKeysForContext(ctx: Context): string[] {
  const mostRestrictiveKey = keyFromDependency(ctx);
  const cachedKeys = contextCacheKey.get(mostRestrictiveKey);
  if (cachedKeys) {
    return cachedKeys;
  }

  const results: string[] = [];
  const n = DIM_ORDER.length;

  // If tableName is undefined in the context, we cannot create combos that "specify" it.
  const tableIdx = DIM_ORDER.indexOf("tableName");

  const totalMasks = 1 << n; // 2^n
  for (let mask = 0; mask < totalMasks; mask++) {
    // Skip combos that specify tableName when ctx.tableName is undefined.
    if (ctx.tableName === undefined && ((mask >> tableIdx) & 1) === 1) continue;

    const dep: ContextDependency = {};
    for (let i = 0; i < n; i++) {
      if (((mask >> i) & 1) === 1) {
        const k = DIM_ORDER[i];
        const v = (ctx as any)[k!];
        // Only assign when we actually have a value (guards tableName: undefined)
        if (v !== undefined) (dep as any)[k!] = v;
      }
    }
    results.push(keyFromDependency(dep));
  }

  contextCacheKey.set(mostRestrictiveKey, results);

  return results;
}

export function getContextDependencyKey(contextDependency: ContextDependency) {
  const keys: (string | number)[] = [];
  contextDependencyKeys.forEach((key) => {
    if (contextDependency[key] !== undefined) {
      keys.push(`${key}:[${contextDependency[key]}]`);
    } else {
      keys.push(`${key}:*`);
    }
  });
  return keys.join(",");
}
