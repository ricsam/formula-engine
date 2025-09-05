import type {
  CellAddress,
  FormulaEngineEvents,
  SerializedCellValue,
  Sheet,
  SpreadsheetRange,
  FiniteSpreadsheetRange,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";
import { renameSheetInFormula } from "../sheet-renamer";

export interface SheetManagerEvents {
  "sheet-added": { sheetName: string };
  "sheet-removed": { sheetName: string };
  "sheet-renamed": { oldName: string; newName: string };
}

export class SheetManager {
  private sheets: Map<string, Sheet> = new Map();
  private eventEmitter?: {
    emit<K extends keyof FormulaEngineEvents>(
      event: K,
      data: FormulaEngineEvents[K]
    ): void;
  };

  constructor(eventEmitter?: {
    emit<K extends keyof FormulaEngineEvents>(
      event: K,
      data: FormulaEngineEvents[K]
    ): void;
  }) {
    this.eventEmitter = eventEmitter;
  }

  getSheets(): Map<string, Sheet> {
    return this.sheets;
  }

  getSheet(name: string): Sheet | undefined {
    return this.sheets.get(name);
  }

  addSheet(name: string): Sheet {
    const sheet = {
      name,
      index: this.sheets.size,
      content: new Map(),
    };

    if (this.sheets.has(sheet.name)) {
      throw new Error("Sheet already exists");
    }

    this.sheets.set(name, sheet);

    // Emit sheet-added event
    this.eventEmitter?.emit("sheet-added", {
      sheetName: name,
    });
    return sheet;
  }

  removeSheet(sheetName: string): Sheet {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Remove the sheet
    this.sheets.delete(sheetName);

    // Emit sheet-removed event
    this.eventEmitter?.emit("sheet-removed", {
      sheetName: sheetName,
    });

    return sheet;
  }

  renameSheet(sheetName: string, newName: string): Sheet {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    if (this.sheets.has(newName)) {
      throw new Error("Sheet with new name already exists");
    }

    // Update sheet name
    sheet.name = newName;

    // Update sheets map
    this.sheets.set(newName, sheet);
    this.sheets.delete(sheetName);

    // Emit sheet-renamed event
    this.eventEmitter?.emit("sheet-renamed", {
      oldName: sheetName,
      newName: newName,
    });

    return sheet;
  }

  updateFormulasForSheetRename(
    oldName: string,
    newName: string,
    updateCallback: (formula: string) => string = (formula) =>
      renameSheetInFormula(formula, oldName, newName)
  ): void {
    // Update all formulas that reference this sheet
    this.sheets.forEach((sheet) => {
      sheet.content.forEach((cell, key) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          const formula = cell.slice(1);
          const updatedFormula = updateCallback(formula);

          // Only update if the formula actually changed
          if (updatedFormula !== formula) {
            sheet.content.set(key, `=${updatedFormula}`);
          }
        }
      });
    });
  }

  getSheetSerialized(sheetName: string): Map<string, SerializedCellValue> {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) return new Map();

    return sheet.content;
  }

  setCellContent(address: CellAddress, content: SerializedCellValue): void {
    const sheet = this.sheets.get(address.sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    sheet.content.set(getCellReference(address), content);
  }

  reevaluateSheet(
    sheetName: string,
    evaluateCallback: (address: CellAddress) => void
  ): void {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    for (const key of sheet.content.keys()) {
      const address = parseCellReference(key);
      evaluateCallback({ ...address, sheetName });
    }
  }

  /**
   * Replace all content for a sheet (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setSheetContent(
    sheetName: string,
    newContent: Map<string, SerializedCellValue>
  ): void {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Clear existing content without breaking the Map reference
    sheet.content.clear();

    // Repopulate with new content
    newContent.forEach((value, key) => {
      sheet.content.set(key, value);
    });

    // Note: No specific sheet-updated event defined, content changes are handled elsewhere
  }

  /**
   * Converts a SpreadsheetRange to FiniteSpreadsheetRange, throwing an error if infinite
   */
  private toFiniteRange(range: SpreadsheetRange): FiniteSpreadsheetRange {
    if (range.end.col.type === "infinity" || range.end.row.type === "infinity") {
      throw new Error("Clearing infinite ranges is not supported");
    }
    
    return {
      start: range.start,
      end: {
        col: range.end.col.value,
        row: range.end.row.value,
      },
    };
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   */
  clearSpreadsheetRange(
    sheetName: string,
    range: SpreadsheetRange,
    setSheetContent: (
      sheetName: string,
      content: Map<string, SerializedCellValue>
    ) => void
  ) {
    // Check if range has infinite ends - not supported for now
    if (
      range.end.col.type === "infinity" ||
      range.end.row.type === "infinity"
    ) {
      throw new Error("Clearing infinite ranges is not supported");
    }

    const sheet = this.sheets.get(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Get current sheet content and prepare new content with cleared cells
    const newContent = new Map(sheet.content);

    // Convert to finite range (throws error if infinite)
    const finiteRange = this.toFiniteRange(range);

    // Iterate through all cells in the range and clear them
    const startCol = finiteRange.start.col;
    const startRow = finiteRange.start.row;
    const endCol = finiteRange.end.col;
    const endRow = finiteRange.end.row;

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cellRef = getCellReference({
          colIndex: col,
          rowIndex: row,
        });
        newContent.delete(cellRef);
      }
    }

    // Update sheet content in a single operation
    setSheetContent(sheetName, newContent);
  }
}
