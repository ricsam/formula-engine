import type { DependencyNode } from "../types";

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
      return `cell:${node.sheetName}:${node.address.rowIndex}:${node.address.colIndex}`;

    case "range":
      const endCol =
        node.range.end.col.type === "number"
          ? node.range.end.col.value.toString()
          : "INFINITY";
      const endRow =
        node.range.end.row.type === "number"
          ? node.range.end.row.value.toString()
          : "INFINITY";
      return `range:${node.sheetName}:${node.range.start.row}:${node.range.start.col}:${endRow}:${endCol}`;

    case "multi-spreadsheet-range":
      const endColMulti =
        node.ranges.end.col.type === "number"
          ? node.ranges.end.col.value.toString()
          : "INFINITY";
      const endRowMulti =
        node.ranges.end.row.type === "number"
          ? node.ranges.end.row.value.toString()
          : "INFINITY";

      if (node.sheetNames.type === "list") {
        const sheetList = node.sheetNames.list.join(",");
        return `multi-range:list:${sheetList}:${node.ranges.start.row}:${node.ranges.start.col}:${endRowMulti}:${endColMulti}`;
      } else {
        return `multi-range:range:${node.sheetNames.startSpreadsheetName}:${node.sheetNames.endSpreadsheetName}:${node.ranges.start.row}:${node.ranges.start.col}:${endRowMulti}:${endColMulti}`;
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
      if (parts.length !== 4) {
        throw new Error(`Invalid cell key format: ${key}`);
      }
      const sheetName = parts[1];
      const rowStr = parts[2];
      const colStr = parts[3];

      if (
        sheetName === undefined ||
        rowStr === undefined ||
        colStr === undefined
      ) {
        throw new Error(`Invalid cell key format: ${key}`);
      }

      return {
        type: "cell",
        sheetName,
        address: {
          rowIndex: parseInt(rowStr, 10),
          colIndex: parseInt(colStr, 10),
        },
      };
    }

    case "range": {
      if (parts.length !== 6) {
        throw new Error(`Invalid range key format: ${key}`);
      }

      const sheetName = parts[1];
      const startRowStr = parts[2];
      const startColStr = parts[3];
      const endRowStr = parts[4];
      const endColStr = parts[5];

      if (
        sheetName === undefined ||
        startRowStr === undefined ||
        startColStr === undefined ||
        endRowStr === undefined ||
        endColStr === undefined
      ) {
        throw new Error(`Invalid range key format: ${key}`);
      }

      const parseEndValue = (value: string) => {
        if (value === "INFINITY") {
          return { type: "infinity" as const, sign: "positive" as const };
        }
        return { type: "number" as const, value: parseInt(value, 10) };
      };

      return {
        type: "range",
        sheetName,
        range: {
          start: {
            row: parseInt(startRowStr, 10),
            col: parseInt(startColStr, 10),
          },
          end: {
            row: parseEndValue(endRowStr),
            col: parseEndValue(endColStr),
          },
        },
      };
    }

    case "multi-range": {
      if (parts.length < 7) {
        throw new Error(`Invalid multi-range key format: ${key}`);
      }

      const parseEndValue = (value: string) => {
        if (value === "INFINITY") {
          return { type: "infinity" as const, sign: "positive" as const };
        }
        return { type: "number" as const, value: parseInt(value, 10) };
      };

      const sheetNamesType = parts[1];

      if (sheetNamesType === "list") {
        if (parts.length !== 7) {
          throw new Error(`Invalid multi-range list key format: ${key}`);
        }

        const sheetListStr = parts[2];
        const startRowStr = parts[3];
        const startColStr = parts[4];
        const endRowStr = parts[5];
        const endColStr = parts[6];

        if (
          sheetListStr === undefined ||
          startRowStr === undefined ||
          startColStr === undefined ||
          endRowStr === undefined ||
          endColStr === undefined
        ) {
          throw new Error(`Invalid multi-range list key format: ${key}`);
        }

        const sheetList = sheetListStr === "" ? [] : sheetListStr.split(",");

        return {
          type: "multi-spreadsheet-range",
          ranges: {
            start: {
              row: parseInt(startRowStr, 10),
              col: parseInt(startColStr, 10),
            },
            end: {
              row: parseEndValue(endRowStr),
              col: parseEndValue(endColStr),
            },
          },
          sheetNames: {
            type: "list",
            list: sheetList,
          },
        };
      } else if (sheetNamesType === "range") {
        if (parts.length !== 8) {
          throw new Error(`Invalid multi-range range key format: ${key}`);
        }

        const startSheetName = parts[2];
        const endSheetName = parts[3];
        const startRowStr = parts[4];
        const startColStr = parts[5];
        const endRowStr = parts[6];
        const endColStr = parts[7];

        if (
          startSheetName === undefined ||
          endSheetName === undefined ||
          startRowStr === undefined ||
          startColStr === undefined ||
          endRowStr === undefined ||
          endColStr === undefined
        ) {
          throw new Error(`Invalid multi-range range key format: ${key}`);
        }

        return {
          type: "multi-spreadsheet-range",
          ranges: {
            start: {
              row: parseInt(startRowStr, 10),
              col: parseInt(startColStr, 10),
            },
            end: {
              row: parseEndValue(endRowStr),
              col: parseEndValue(endColStr),
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

        if (
          !["Headers", "All", "AllData"].includes(capitalizedAreaType)
        ) {
          throw new Error(`Invalid table area type: ${areaType}`);
        }

        return {
          type: "table",
          tableName,
          sheetName,
          area: {
            kind: capitalizedAreaType as
              | "Headers"
              | "All"
              | "AllData",
          },
        };
      }
    }

    default:
      throw new Error(`Unknown dependency node type: ${type}`);
  }
}
