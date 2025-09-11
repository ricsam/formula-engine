import type {
  CellAddress,
  FormulaEngineEvents,
  SerializedCellValue,
  Sheet,
  SpreadsheetRange,
  FiniteSpreadsheetRange,
  Workbook,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";
import { renameSheetInFormula } from "../sheet-renamer";

export class WorkbookManager {
  private workbooks: Map<string, Workbook> = new Map();
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

  getSheets(workbookName: string): Map<string, Sheet> {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }
    return workbook.sheets;
  }

  getWorkbooks(): Map<string, Workbook> {
    return this.workbooks;
  }

  getSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Sheet | undefined {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }
    return sheet;
  }

  addSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Sheet {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }
    const sheet = {
      name: sheetName,
      index: workbook.sheets.size,
      content: new Map(),
    };

    if (workbook.sheets.has(sheet.name)) {
      throw new Error("Sheet already exists");
    }

    workbook.sheets.set(sheetName, sheet);

    // Emit sheet-added event
    this.eventEmitter?.emit("sheet-added", {
      sheetName: sheetName,
      workbookName: workbookName,
    });
    return sheet;
  }

  removeSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Sheet {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Remove the sheet
    workbook.sheets.delete(sheetName);

    // Emit sheet-removed event
    this.eventEmitter?.emit("sheet-removed", {
      sheetName: sheetName,
      workbookName: workbookName,
    });

    return sheet;
  }

  renameSheet({
    workbookName,
    sheetName,
    newSheetName,
  }: {
    workbookName: string;
    sheetName: string;
    newSheetName: string;
  }): Sheet {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    if (workbook.sheets.has(newSheetName)) {
      throw new Error("Sheet with new name already exists");
    }

    // Update sheet name
    sheet.name = newSheetName;

    // Update sheets map
    workbook.sheets.set(newSheetName, sheet);
    workbook.sheets.delete(sheetName);

    // Emit sheet-renamed event
    this.eventEmitter?.emit("sheet-renamed", {
      oldSheetName: sheetName,
      newSheetName: newSheetName,
      workbookName: workbookName,
    });

    return sheet;
  }

  updateAllFormulas(updateCallback: (formula: string) => string): void {
    const update = (map: Map<string, Sheet>) => {
      map.forEach((sheet) => {
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
    };

    this.workbooks.forEach((workbook) => {
      update(workbook.sheets);
    });
  }

  getSheetSerialized({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Map<string, SerializedCellValue> {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    return sheet.content;
  }

  setCellContent(address: CellAddress, content: SerializedCellValue): void {
    const sheet = this.getSheet({
      sheetName: address.sheetName,
      workbookName: address.workbookName,
    });

    if (!sheet) {
      throw new Error("Sheet not found");
    }

    sheet.content.set(getCellReference(address), content);
  }

  reevaluateSheet(
    opts: { sheetName: string; workbookName: string },
    evaluateCallback: (address: CellAddress) => void
  ): void {
    const sheet = this.getSheet(opts);

    if (!sheet) {
      throw new Error("Sheet not found");
    }

    for (const key of sheet.content.keys()) {
      const address = parseCellReference(key);
      evaluateCallback({
        ...address,
        sheetName: opts.sheetName,
        workbookName: opts.workbookName,
      });
    }
  }

  /**
   * Replace all content for a sheet (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setSheetContent(
    opts: { sheetName: string; workbookName: string },
    newContent: Map<string, SerializedCellValue>
  ): void {
    const sheet = this.getSheet(opts);
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
    if (
      range.end.col.type === "infinity" ||
      range.end.row.type === "infinity"
    ) {
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
    opts: { sheetName: string; workbookName: string },
    range: SpreadsheetRange,
    setSheetContent: (content: Map<string, SerializedCellValue>) => void
  ) {
    // Check if range has infinite ends - not supported for now
    if (
      range.end.col.type === "infinity" ||
      range.end.row.type === "infinity"
    ) {
      throw new Error("Clearing infinite ranges is not supported");
    }

    const sheet = this.getSheet(opts);

    if (!sheet) {
      throw new Error("Sheet not found");
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
    setSheetContent(newContent);
  }
}
