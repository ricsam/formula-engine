/**
 * Core type definitions for FormulaEngine
 * This file contains all fundamental types used throughout the engine
 */

// Cell addressing types
export interface CellAddress {
  sheetName: string;
  col: number;
  row: number;
}

export type CellNumber =
  | {
      type: "number";
      value: number;
    }
  | {
      type: "infinity";
    };

export type SpreadsheetRange = {
  start: {
    col: number;
    row: number;
  };
  end: {
    col: CellNumber;
    row: CellNumber;
  };
};

export type CellString = {
  type: "string";
  value: string;
};

export type CellBoolean = {
  type: "boolean";
  value: boolean;
};

// Cell value types
export type CellValue = CellNumber | CellString | CellBoolean;
export type SerializedCellValue = string | number | boolean | undefined;

// Named expressions
export interface NamedExpression {
  id: string; // unique identifier for the named expression
  name: string;
  expression: string;
  sheetName?: string; // undefined = global scope, string = sheet name
}

// Formula errors
export enum FormulaError {
  DIV0 = "#DIV/0!",
  NA = "#N/A",
  NAME = "#NAME?",
  NUM = "#NUM!",
  REF = "#REF!",
  VALUE = "#VALUE!",
  CYCLE = "#CYCLE!",
  ERROR = "#ERROR!",
  SPILL = "#SPILL!",
}

// Internal cell structure
export interface Cell {
  value: CellValue;
  formula?: string;
  dependencies?: Set<string>;
  dependents?: Set<string>;
  spilledCells?: SpreadsheetRange;
}

// Sheet structure
export interface Sheet {
  name: string;
  index: number; // 0-based index of the sheet
  cells: Map<string, Cell>;
}

// Parser and AST types
export interface ASTNode {
  type: "function" | "reference" | "value" | "operator" | "array" | "range";
  value?: any;
  children?: ASTNode[];
  functionName?: string;
  operator?: string;
  reference?: CellAddress | SpreadsheetRange;
}

// Event types
export interface FormulaEngineEvents {
  "sheet-added": {
    sheetName: string;
  };
  "sheet-removed": {
    sheetName: string;
  };
  "sheet-renamed": {
    oldName: string;
    newName: string;
  };
}

// Utility types
type AddressKey = string; // Format: "sheet:col:row" e.g., "0:1:2"

// Helper function to create address keys
export function addressToKey(address: CellAddress): AddressKey {
  return `${address.sheet}:${address.col}:${address.row}`;
}

// Helper function to parse address keys
export function keyToAddress(key: AddressKey): CellAddress {
  const parts = key.split(":").map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) {
    throw new Error(`Invalid address key: ${key}`);
  }
  const [sheet, col, row] = parts;
  return { sheet: sheet!, col: col!, row: row! };
}

// Helper function to parse range keys (format: "sheet:startCol:startRow:endCol:endRow")
export function keyToRange(key: string): SimpleCellRange | null {
  const parts = key.split(":").map(Number);
  if (parts.length !== 5 || parts.some(isNaN)) {
    return null;
  }
  const [sheet, startCol, startRow, endCol, endRow] = parts;
  return {
    start: { sheet: sheet!, col: startCol!, row: startRow! },
    end: { sheet: sheet!, col: endCol!, row: endRow! },
  };
}

// Helper function to parse dependency keys - returns type and parsed object
export function parseDependencyKey(
  key: string
):
  | { type: "cell"; address: CellAddress }
  | { type: "range"; range: SimpleCellRange }
  | { type: "named"; name: string; scope?: number } {
  if (key.startsWith("name:")) {
    // Named expression: name:MyValue or name:0:MyValue
    const nameParts = key.split(":");
    if (nameParts.length === 2 && nameParts[1]) {
      return { type: "named", name: nameParts[1] };
    } else if (nameParts.length === 3 && nameParts[1] && nameParts[2]) {
      return {
        type: "named",
        name: nameParts[2],
        scope: parseInt(nameParts[1]),
      };
    }
    throw new Error(`Invalid named expression dependency key: ${key}`);
  }

  const parts = key.split(":");
  if (parts.length === 3 && parts.every((p) => !isNaN(Number(p)))) {
    // Cell dependency: sheet:col:row
    try {
      const address = keyToAddress(key as AddressKey);
      return { type: "cell", address };
    } catch (error) {
      throw new Error(`Invalid cell dependency key: ${key}. ${error}`);
    }
  } else if (parts.length === 5 && parts.every((p) => !isNaN(Number(p)))) {
    // Range dependency: sheet:startCol:startRow:endCol:endRow
    const range = keyToRange(key);
    if (range) {
      return { type: "range", range };
    }
    throw new Error(`Invalid range dependency key: ${key}`);
  }

  throw new Error(`Unknown dependency key format: ${key}`);
}

// A1 notation helpers
export function colNumberToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export function letterToColNumber(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result - 1;
}
