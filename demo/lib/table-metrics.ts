import type { TableDefinition } from "../../src/core/types";

export function getFiniteTableLastRowIndex(
  table: TableDefinition
): number | undefined {
  return table.endRow.type === "number" ? table.endRow.value : undefined;
}

export function getFiniteTableRowCount(
  table: TableDefinition
): number | undefined {
  const lastRowIndex = getFiniteTableLastRowIndex(table);
  if (lastRowIndex === undefined) {
    return undefined;
  }

  return lastRowIndex - table.start.rowIndex + 1;
}

export function isTableLastRow(
  table: TableDefinition,
  rowIndex: number
): boolean {
  const lastRowIndex = getFiniteTableLastRowIndex(table);
  return lastRowIndex !== undefined && rowIndex === lastRowIndex;
}
