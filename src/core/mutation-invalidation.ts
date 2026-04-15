import type { WorkbookManager } from "./managers/workbook-manager";
import type {
  CellAddress,
  RangeAddress,
  SerializedCellValue,
  TableDefinition,
} from "./types";
import { getNamedExpressionResourceKey } from "./resource-keys";
import { getCellReference, parseCellReference } from "./utils";

type CellContentKind = "empty" | "scalar" | "formula";

export type RemovedScope =
  | { type: "workbook"; workbookName: string }
  | { type: "sheet"; workbookName: string; sheetName: string };

export type MutationInvalidation = {
  touchedCells: Array<{
    address: CellAddress;
    beforeKind: CellContentKind;
    afterKind: CellContentKind;
  }>;
  /**
   * Cells whose table membership or implicit current-row table context changed
   * without necessarily changing their formula text.
   */
  tableContextChangedCells?: CellAddress[];
  resourceKeys: string[];
  removedScopes?: RemovedScope[];
};

function getSerializedCellValueKind(
  value: SerializedCellValue | undefined
): CellContentKind {
  if (
    value === undefined ||
    (typeof value === "string" && value.length === 0)
  ) {
    return "empty";
  }
  if (typeof value === "string" && value.startsWith("=")) {
    return "formula";
  }
  return "scalar";
}

function emptyMutationInvalidation(): MutationInvalidation {
  return {
    touchedCells: [],
    resourceKeys: [],
  };
}

export function getMutationAddressKey(address: CellAddress): string {
  return `${address.workbookName}:${address.sheetName}:${getCellReference(
    address
  )}`;
}

export function buildTouchedCells(
  cells: Array<{
    address: CellAddress;
    before: SerializedCellValue | undefined;
    after: SerializedCellValue | undefined;
  }>
): MutationInvalidation["touchedCells"] {
  const deduped = new Map<
    string,
    MutationInvalidation["touchedCells"][number]
  >();

  for (const cell of cells) {
    deduped.set(getMutationAddressKey(cell.address), {
      address: cell.address,
      beforeKind: getSerializedCellValueKind(cell.before),
      afterKind: getSerializedCellValueKind(cell.after),
    });
  }

  return Array.from(deduped.values());
}

export function buildFormulaTouchedCells(
  cells: CellAddress[]
): MutationInvalidation["touchedCells"] {
  return cells.map((address) => ({
    address,
    beforeKind: "formula",
    afterKind: "formula",
  }));
}

export function captureCellContents(
  workbookManager: WorkbookManager,
  addresses: CellAddress[]
): Map<string, SerializedCellValue | undefined> {
  const contents = new Map<string, SerializedCellValue | undefined>();
  for (const address of addresses) {
    try {
      contents.set(
        getMutationAddressKey(address),
        workbookManager.getCellContent(address)
      );
    } catch {
      contents.set(getMutationAddressKey(address), undefined);
    }
  }
  return contents;
}

export function buildSheetContentTouchedCells(
  opts: { workbookName: string; sheetName: string },
  beforeContent: Map<string, SerializedCellValue> | undefined,
  afterContent: Map<string, SerializedCellValue>
): MutationInvalidation["touchedCells"] {
  const touchedKeys = new Set<string>([
    ...Array.from(beforeContent?.keys() ?? []),
    ...Array.from(afterContent.keys()),
  ]);

  return buildTouchedCells(
    Array.from(touchedKeys, (key) => ({
      address: {
        workbookName: opts.workbookName,
        sheetName: opts.sheetName,
        ...parseCellReference(key),
      },
      before: beforeContent?.get(key),
      after: afterContent.get(key),
    }))
  );
}

function getTableCellKey(address: CellAddress): string {
  return `${address.workbookName}:${address.sheetName}:${address.rowIndex}:${address.colIndex}`;
}

function collectTableFootprintCells(
  workbookManager: WorkbookManager,
  table: TableDefinition
): Array<{
  address: CellAddress;
  content: SerializedCellValue | undefined;
}> {
  const cells = new Map<
    string,
    {
      address: CellAddress;
      content: SerializedCellValue | undefined;
    }
  >();
  const sheet = workbookManager.getSheet({
    workbookName: table.workbookName,
    sheetName: table.sheetName,
  });
  if (!sheet) {
    return [];
  }

  const startColIndex = table.start.colIndex;
  const endColIndex = startColIndex + table.headers.size - 1;

  if (table.endRow.type === "number") {
    for (
      let rowIndex = table.start.rowIndex;
      rowIndex <= table.endRow.value;
      rowIndex++
    ) {
      for (
        let colIndex = startColIndex;
        colIndex <= endColIndex;
        colIndex++
      ) {
        const address = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          rowIndex,
          colIndex,
        };
        cells.set(getTableCellKey(address), {
          address,
          content: workbookManager.getCellContent(address),
        });
      }
    }
    return Array.from(cells.values());
  }

  for (const [ref, content] of sheet.content.entries()) {
    const { rowIndex, colIndex } = parseCellReference(ref);
    if (rowIndex < table.start.rowIndex) {
      continue;
    }
    if (colIndex < startColIndex || colIndex > endColIndex) {
      continue;
    }

    const address = {
      workbookName: table.workbookName,
      sheetName: table.sheetName,
      rowIndex,
      colIndex,
    };
    cells.set(getTableCellKey(address), {
      address,
      content,
    });
  }

  return Array.from(cells.values());
}

export function buildTableTouchedCells(
  workbookManager: WorkbookManager,
  tables: Array<TableDefinition | undefined>
): MutationInvalidation["touchedCells"] {
  const touchedCells = new Map<
    string,
    MutationInvalidation["touchedCells"][number]
  >();

  for (const table of tables) {
    if (!table) {
      continue;
    }
    for (const cell of collectTableFootprintCells(workbookManager, table)) {
      touchedCells.set(getTableCellKey(cell.address), {
        address: cell.address,
        beforeKind: getSerializedCellValueKind(cell.content),
        afterKind: getSerializedCellValueKind(cell.content),
      });
    }
  }

  return Array.from(touchedCells.values());
}

export function buildTableContextChangedCells(
  workbookManager: WorkbookManager,
  tables: Array<TableDefinition | undefined>
): CellAddress[] {
  const changedCells = new Map<string, CellAddress>();

  for (const table of tables) {
    if (!table) {
      continue;
    }
    for (const cell of collectTableFootprintCells(workbookManager, table)) {
      changedCells.set(getTableCellKey(cell.address), cell.address);
    }
  }

  return Array.from(changedCells.values());
}

export function mergeTouchedCells(
  ...groups: MutationInvalidation["touchedCells"][]
): MutationInvalidation["touchedCells"] {
  const precedence = {
    empty: 0,
    scalar: 1,
    formula: 2,
  } as const;
  const merged = new Map<
    string,
    MutationInvalidation["touchedCells"][number]
  >();

  for (const group of groups) {
    for (const touchedCell of group) {
      const key = getTableCellKey(touchedCell.address);
      const existing = merged.get(key);
      if (!existing) {
        merged.set(key, touchedCell);
        continue;
      }

      merged.set(key, {
        address: touchedCell.address,
        beforeKind:
          precedence[touchedCell.beforeKind] >= precedence[existing.beforeKind]
            ? touchedCell.beforeKind
            : existing.beforeKind,
        afterKind:
          precedence[touchedCell.afterKind] >= precedence[existing.afterKind]
            ? touchedCell.afterKind
            : existing.afterKind,
      });
    }
  }

  return Array.from(merged.values());
}

export function getNamedExpressionScopeResourceKeys(
  expressions: Iterable<string>,
  opts: {
    workbookName?: string;
    sheetName?: string;
  }
): string[] {
  return Array.from(
    new Set(
      Array.from(expressions, (expressionName) =>
        getNamedExpressionResourceKey({
          expressionName,
          workbookName: opts.workbookName,
          sheetName: opts.sheetName,
        })
      )
    )
  );
}

export function getFiniteRangeAddresses(address: RangeAddress): CellAddress[] {
  if (
    address.range.end.col.type === "infinity" ||
    address.range.end.row.type === "infinity"
  ) {
    return [];
  }

  const cells: CellAddress[] = [];
  for (
    let colIndex = address.range.start.col;
    colIndex <= address.range.end.col.value;
    colIndex++
  ) {
    for (
      let rowIndex = address.range.start.row;
      rowIndex <= address.range.end.row.value;
      rowIndex++
    ) {
      cells.push({
        workbookName: address.workbookName,
        sheetName: address.sheetName,
        colIndex,
        rowIndex,
      });
    }
  }
  return cells;
}
