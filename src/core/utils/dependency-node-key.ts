import type { DependencyNode } from "../types";
import { getRowNumber, indexToColumn, parseCellReference, columnToIndex } from "../utils";

export function dependencyNodeToKey(node: DependencyNode): string {
  switch (node.type) {
    case "cell":
      if (
        node.address.rowIndex === undefined ||
        node.address.colIndex === undefined
      ) {
        throw new Error(
          `Invalid cell address: rowIndex and colIndex must be defined (got rowIndex=${node.address.rowIndex}, colIndex=${node.address.colIndex})`
        );
      }
      const cellRef = `${indexToColumn(node.address.colIndex)}${getRowNumber(node.address.rowIndex)}`;
      return `cell:${node.sheetName}:${cellRef}`;

    case "range":
      const startCell = `${indexToColumn(node.range.start.col)}${getRowNumber(node.range.start.row)}`;
      
      // Handle different range types according to canonical format
      let rangeEnd: string;
      if (node.range.end.row.type === "infinity" && node.range.end.col.type === "infinity") {
        // Both infinite: A5:INFINITY
        rangeEnd = "INFINITY";
      } else if (node.range.end.row.type === "infinity" && node.range.end.col.type === "number") {
        // Row infinite, col finite: A5:D (column only)
        rangeEnd = indexToColumn(node.range.end.col.value);
      } else if (node.range.end.row.type === "number" && node.range.end.col.type === "infinity") {
        // Row finite, col infinite: A5:10 (row only)
        rangeEnd = getRowNumber(node.range.end.row.value).toString();
      } else if (node.range.end.row.type === "number" && node.range.end.col.type === "number") {
        // Both finite: A5:D10
        rangeEnd = `${indexToColumn(node.range.end.col.value)}${getRowNumber(node.range.end.row.value)}`;
      } else {
        throw new Error("Invalid range end configuration");
      }
      
      return `range:${node.sheetName}:${startCell}:${rangeEnd}`;

    case "multi-spreadsheet-range":
      const startCellMulti = `${indexToColumn(node.ranges.start.col)}${getRowNumber(node.ranges.start.row)}`;
      
      // Handle different range types according to canonical format
      let rangeEndMulti: string;
      if (node.ranges.end.row.type === "infinity" && node.ranges.end.col.type === "infinity") {
        // Both infinite: A5:INFINITY
        rangeEndMulti = "INFINITY";
      } else if (node.ranges.end.row.type === "infinity" && node.ranges.end.col.type === "number") {
        // Row infinite, col finite: A5:D (column only)
        rangeEndMulti = indexToColumn(node.ranges.end.col.value);
      } else if (node.ranges.end.row.type === "number" && node.ranges.end.col.type === "infinity") {
        // Row finite, col infinite: A5:10 (row only)
        rangeEndMulti = getRowNumber(node.ranges.end.row.value).toString();
      } else if (node.ranges.end.row.type === "number" && node.ranges.end.col.type === "number") {
        // Both finite: A5:D10
        rangeEndMulti = `${indexToColumn(node.ranges.end.col.value)}${getRowNumber(node.ranges.end.row.value)}`;
      } else {
        throw new Error("Invalid range end configuration");
      }

      if (node.sheetNames.type === "list") {
        const sheetList = node.sheetNames.list.join(",");
        return `multi-range:list:${sheetList}:${startCellMulti}:${rangeEndMulti}`;
      } else {
        return `multi-range:range:${node.sheetNames.startSpreadsheetName}:${node.sheetNames.endSpreadsheetName}:${startCellMulti}:${rangeEndMulti}`;
      }

    case "named-expression":
      if (node.sheetName) {
        return `named:${node.sheetName}:${node.name}`;
      } else {
        return `named:global:${node.name}`;
      }

    case "table":
      if (node.area.kind === "Data") {
        const columns = node.area.columns.join(",");
        return `table:${node.sheetName}:${node.tableName}:data:${columns}`;
      } else {
        return `table:${node.sheetName}:${node.tableName}:${node.area.kind}`;
      }

    default:
      throw new Error(`Unknown dependency node type: ${(node as any).type}`);
  }
}

export function keyToDependencyNode(key: string): DependencyNode {
  const parts = key.split(":");

  if (parts.length < 2) {
    throw new Error(`Invalid dependency key format: ${key}`);
  }

  const type = parts[0];

  switch (type) {
    case "cell": {
      if (parts.length !== 3) {
        throw new Error(`Invalid cell key format: ${key}`);
      }
      const sheetName = parts[1];
      const cellRef = parts[2];

      if (sheetName === undefined || cellRef === undefined) {
        throw new Error(`Invalid cell key format: ${key}`);
      }

      const { rowIndex, colIndex } = parseCellReference(cellRef);

      return {
        type: "cell",
        sheetName,
        address: {
          rowIndex,
          colIndex,
        },
      };
    }

    case "range": {
      if (parts.length !== 4) {
        throw new Error(`Invalid range key format: ${key}`);
      }

      const sheetName = parts[1];
      const startCellRef = parts[2];
      const endRef = parts[3];

      if (
        sheetName === undefined ||
        startCellRef === undefined ||
        endRef === undefined
      ) {
        throw new Error(`Invalid range key format: ${key}`);
      }

      // Parse start cell
      const { rowIndex: startRow, colIndex: startCol } = parseCellReference(startCellRef);

      // Parse end reference - could be INFINITY, column only, row only, or full cell
      let endRow: { type: "number"; value: number } | { type: "infinity"; sign: "positive" };
      let endCol: { type: "number"; value: number } | { type: "infinity"; sign: "positive" };

      if (endRef === "INFINITY") {
        // Both infinite: A5:INFINITY
        endRow = { type: "infinity", sign: "positive" };
        endCol = { type: "infinity", sign: "positive" };
      } else if (/^[A-Z]+$/.test(endRef)) {
        // Column only: A5:D (row infinite, col finite)
        endRow = { type: "infinity", sign: "positive" };
        endCol = { type: "number", value: columnToIndex(endRef) };
      } else if (/^\d+$/.test(endRef)) {
        // Row only: A5:10 (row finite, col infinite)
        endRow = { type: "number", value: parseInt(endRef, 10) - 1 }; // Convert to 0-based
        endCol = { type: "infinity", sign: "positive" };
      } else {
        // Full cell: A5:D10 (both finite)
        const { rowIndex: endRowIndex, colIndex: endColIndex } = parseCellReference(endRef);
        endRow = { type: "number", value: endRowIndex };
        endCol = { type: "number", value: endColIndex };
      }

      return {
        type: "range",
        sheetName,
        range: {
          start: {
            row: startRow,
            col: startCol,
          },
          end: {
            row: endRow,
            col: endCol,
          },
        },
      };
    }

    case "multi-range": {
      if (parts.length < 5) {
        throw new Error(`Invalid multi-range key format: ${key}`);
      }

      const sheetNamesType = parts[1];

      if (sheetNamesType === "list") {
        if (parts.length !== 5) {
          throw new Error(`Invalid multi-range list key format: ${key}`);
        }

        const sheetListStr = parts[2];
        const startCellRef = parts[3];
        const endRef = parts[4];

        if (
          sheetListStr === undefined ||
          startCellRef === undefined ||
          endRef === undefined
        ) {
          throw new Error(`Invalid multi-range list key format: ${key}`);
        }

        const sheetList = sheetListStr === "" ? [] : sheetListStr.split(",");

        // Parse start cell
        const { rowIndex: startRow, colIndex: startCol } = parseCellReference(startCellRef);

        // Parse end reference - same logic as regular ranges
        let endRow: { type: "number"; value: number } | { type: "infinity"; sign: "positive" };
        let endCol: { type: "number"; value: number } | { type: "infinity"; sign: "positive" };

        if (endRef === "INFINITY") {
          endRow = { type: "infinity", sign: "positive" };
          endCol = { type: "infinity", sign: "positive" };
        } else if (/^[A-Z]+$/.test(endRef)) {
          endRow = { type: "infinity", sign: "positive" };
          endCol = { type: "number", value: columnToIndex(endRef) };
        } else if (/^\d+$/.test(endRef)) {
          endRow = { type: "number", value: parseInt(endRef, 10) - 1 };
          endCol = { type: "infinity", sign: "positive" };
        } else {
          const { rowIndex: endRowIndex, colIndex: endColIndex } = parseCellReference(endRef);
          endRow = { type: "number", value: endRowIndex };
          endCol = { type: "number", value: endColIndex };
        }

        return {
          type: "multi-spreadsheet-range",
          ranges: {
            start: {
              row: startRow,
              col: startCol,
            },
            end: {
              row: endRow,
              col: endCol,
            },
          },
          sheetNames: {
            type: "list",
            list: sheetList,
          },
        };
      } else if (sheetNamesType === "range") {
        if (parts.length !== 6) {
          throw new Error(`Invalid multi-range range key format: ${key}`);
        }

        const startSheetName = parts[2];
        const endSheetName = parts[3];
        const startCellRef = parts[4];
        const endRef = parts[5];

        if (
          startSheetName === undefined ||
          endSheetName === undefined ||
          startCellRef === undefined ||
          endRef === undefined
        ) {
          throw new Error(`Invalid multi-range range key format: ${key}`);
        }

        // Parse start cell
        const { rowIndex: startRow, colIndex: startCol } = parseCellReference(startCellRef);

        // Parse end reference - same logic as regular ranges
        let endRow: { type: "number"; value: number } | { type: "infinity"; sign: "positive" };
        let endCol: { type: "number"; value: number } | { type: "infinity"; sign: "positive" };

        if (endRef === "INFINITY") {
          endRow = { type: "infinity", sign: "positive" };
          endCol = { type: "infinity", sign: "positive" };
        } else if (/^[A-Z]+$/.test(endRef)) {
          endRow = { type: "infinity", sign: "positive" };
          endCol = { type: "number", value: columnToIndex(endRef) };
        } else if (/^\d+$/.test(endRef)) {
          endRow = { type: "number", value: parseInt(endRef, 10) - 1 };
          endCol = { type: "infinity", sign: "positive" };
        } else {
          const { rowIndex: endRowIndex, colIndex: endColIndex } = parseCellReference(endRef);
          endRow = { type: "number", value: endRowIndex };
          endCol = { type: "number", value: endColIndex };
        }

        return {
          type: "multi-spreadsheet-range",
          ranges: {
            start: {
              row: startRow,
              col: startCol,
            },
            end: {
              row: endRow,
              col: endCol,
            },
          },
          sheetNames: {
            type: "range",
            startSpreadsheetName: startSheetName,
            endSpreadsheetName: endSheetName,
          },
        };
      } else {
        throw new Error(
          `Invalid multi-range sheet names type: ${sheetNamesType}`
        );
      }
    }

    case "named": {
      if (parts.length !== 3) {
        throw new Error(`Invalid named expression key format: ${key}`);
      }

      const scope = parts[1];
      const name = parts[2];

      if (scope === undefined || name === undefined) {
        throw new Error(`Invalid named expression key format: ${key}`);
      }

      return {
        type: "named-expression",
        name,
        sheetName: scope,
      };
    }

    case "table": {
      if (parts.length < 4) {
        throw new Error(`Invalid table key format: ${key}`);
      }

      const sheetName = parts[1];
      const tableName = parts[2];
      const areaType = parts[3];

      if (
        sheetName === undefined ||
        tableName === undefined ||
        areaType === undefined
      ) {
        throw new Error(`Invalid table key format: ${key}`);
      }

      if (areaType === "data") {
        if (parts.length !== 5) {
          throw new Error(`Invalid table data key format: ${key}`);
        }
        const columnsStr = parts[4];
        if (columnsStr === undefined) {
          throw new Error(`Invalid table data key format: ${key}`);
        }
        const columns = columnsStr === "" ? [] : columnsStr.split(",");
        return {
          type: "table",
          tableName,
          sheetName,
          area: {
            kind: "Data",
            columns,
            isCurrentRow: false,
          },
        };
      } else {
        if (parts.length !== 4) {
          throw new Error(`Invalid table key format: ${key}`);
        }

        const capitalizedAreaType =
          areaType.charAt(0).toUpperCase() + areaType.slice(1);

        if (!["Headers", "All", "AllData"].includes(capitalizedAreaType)) {
          throw new Error(`Invalid table area type: ${areaType}`);
        }

        return {
          type: "table",
          tableName,
          sheetName,
          area: {
            kind: capitalizedAreaType as "Headers" | "All" | "AllData",
          },
        };
      }
    }

    default:
      throw new Error(`Unknown dependency node type: ${type}`);
  }
}
