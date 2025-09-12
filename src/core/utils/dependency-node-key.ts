import type { DependencyNode } from "../types";
import { getRowNumber, indexToColumn, parseCellReference } from "../utils";

export function dependencyNodeToKey(node: DependencyNode): string {
  if (
    node.address.rowIndex === undefined ||
    node.address.colIndex === undefined
  ) {
    throw new Error(
      `Invalid cell address: rowIndex and colIndex must be defined (got rowIndex=${node.address.rowIndex}, colIndex=${node.address.colIndex})`
    );
  }
  const cellRef = `${indexToColumn(node.address.colIndex)}${getRowNumber(node.address.rowIndex)}`;
  return `cell:${node.workbookName}:${node.sheetName}:${cellRef}`;
}

export function keyToDependencyNode(key: string): DependencyNode {
  const parts = key.split(":");

  if (parts.length < 2) {
    throw new Error(`Invalid dependency key format: ${key}`);
  }

  if (parts.length !== 4) {
    throw new Error(`Invalid cell key format: ${key}`);
  }
  const workbookName = parts[1];
  const sheetName = parts[2];
  const cellRef = parts[3];

  if (
    workbookName === undefined ||
    sheetName === undefined ||
    cellRef === undefined
  ) {
    throw new Error(`Invalid cell key format: ${key}`);
  }

  const { rowIndex, colIndex } = parseCellReference(cellRef);

  if (
    rowIndex === undefined ||
    colIndex === undefined ||
    Number.isNaN(rowIndex) ||
    Number.isNaN(colIndex)
  ) {
    throw new Error(`Invalid cell reference: ${cellRef}`);
  }

  return {
    workbookName,
    sheetName,
    address: {
      rowIndex,
      colIndex,
    },
  };
}
