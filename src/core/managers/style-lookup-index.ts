import type { CellAddress, RangeAddress } from "../types";

export interface StyleAreaMatch<T> {
  readonly rule: T;
  readonly area: RangeAddress;
  readonly ruleIndex: number;
  readonly areaIndex: number;
}

type Entry<T> = StyleAreaMatch<T> & Bounds;
interface Bounds {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
}
interface Node<T> extends Bounds {
  entries?: Entry<T>[];
  left?: Node<T>;
  right?: Node<T>;
}

/** Static, lazily rebuilt per-sheet rectangle tree. Open ranges are never expanded
 * into cells. Matches retain rule AND area order (important for gradients).
 */
export class StyleLookupIndex<T extends { areas: RangeAddress[] }> {
  private readonly sheets = new Map<string, Map<string, Node<T>>>();

  constructor(rules: readonly T[]) {
    const sheets = new Map<string, Map<string, Entry<T>[]>>();
    rules.forEach((rule, ruleIndex) => {
      rule.areas.forEach((area, areaIndex) => {
        let workbook = sheets.get(area.workbookName);
        if (!workbook) sheets.set(area.workbookName, (workbook = new Map()));
        let entries = workbook.get(area.sheetName);
        if (!entries) workbook.set(area.sheetName, (entries = []));
        entries.push({
          rule, area, ruleIndex, areaIndex,
          minRow: area.range.start.row,
          minCol: area.range.start.col,
          maxRow: area.range.end.row.type === "number" ? area.range.end.row.value : Infinity,
          maxCol: area.range.end.col.type === "number" ? area.range.end.col.value : Infinity,
        });
      });
    });
    for (const [name, workbook] of sheets) {
      const trees = new Map<string, Node<T>>();
      for (const [sheet, entries] of workbook) trees.set(sheet, this.build(entries, 0));
      this.sheets.set(name, trees);
    }
  }

  private build(entries: Entry<T>[], depth: number): Node<T> {
    const node: Node<T> = { minRow: Infinity, minCol: Infinity, maxRow: -Infinity, maxCol: -Infinity };
    for (const entry of entries) {
      node.minRow = Math.min(node.minRow, entry.minRow);
      node.minCol = Math.min(node.minCol, entry.minCol);
      node.maxRow = Math.max(node.maxRow, entry.maxRow);
      node.maxCol = Math.max(node.maxCol, entry.maxCol);
    }
    if (entries.length <= 8) {
      node.entries = entries;
    } else {
      const axis = depth % 2 === 0 ? "minRow" : "minCol";
      entries.sort((a, b) => a[axis] - b[axis]);
      const middle = entries.length >>> 1;
      node.left = this.build(entries.slice(0, middle), depth + 1);
      node.right = this.build(entries.slice(middle), depth + 1);
    }
    return node;
  }

  query(cell: CellAddress): StyleAreaMatch<T>[] {
    const matches: Entry<T>[] = [];
    const contains = (bounds: Bounds) => cell.rowIndex >= bounds.minRow &&
      cell.rowIndex <= bounds.maxRow && cell.colIndex >= bounds.minCol &&
      cell.colIndex <= bounds.maxCol;
    const visit = (node: Node<T> | undefined): void => {
      if (!node || !contains(node)) return;
      if (node.entries) {
        for (const entry of node.entries) if (contains(entry)) matches.push(entry);
      } else {
        visit(node.left);
        visit(node.right);
      }
    };
    visit(this.sheets.get(cell.workbookName)?.get(cell.sheetName));
    matches.sort((a, b) => a.ruleIndex - b.ruleIndex || a.areaIndex - b.areaIndex);
    return matches;
  }
}
