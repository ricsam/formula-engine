import { FormulaEngine } from "../engine";

export interface SpreadsheetVisualizerOptions {
  /** Number of rows to display (1-based) */
  numRows: number;
  /** Number of columns to display (0-based, so 11 = A-K) */
  numCols: number;
  /** Starting row (1-based, default: 1) */
  startRow?: number;
  /** Starting column (0-based, default: 0 = A) */
  startCol?: number;
  /** Sheet name to visualize (default: uses engine's current sheet) */
  sheetName?: string;
  /** Character to use for empty cells (default: ".") */
  emptyCellChar?: string;
  /** Minimum column width (default: 3) */
  minColWidth?: number;
  /** Maximum column width (default: 20) */
  maxColWidth?: number;
  /** Whether to show column headers (A, B, C...) */
  showColumnHeaders?: boolean;
  /** Whether to show row numbers */
  showRowNumbers?: boolean;
}

/**
 * Generates a visual ASCII representation of a spreadsheet region
 */
export function visualizeSpreadsheet(
  engine: FormulaEngine,
  options: SpreadsheetVisualizerOptions
): string {
  const {
    numRows,
    numCols,
    startRow = 1,
    startCol = 0,
    sheetName,
    emptyCellChar = ".",
    minColWidth = 3,
    maxColWidth = 20,
    showColumnHeaders = true,
    showRowNumbers = true,
  } = options;

  // Determine the sheet name to use upfront
  const targetSheetName = sheetName || (() => {
    const sheets = engine.sheets;
    const sheetNames = Array.from(sheets.keys());
    return sheetNames.length > 0 ? sheetNames[0]! : null;
  })();

  // First pass: collect all cell values and calculate column widths
  const cellValues: string[][] = [];
  const columnWidths: number[] = new Array(numCols).fill(minColWidth);

  // Collect column headers if needed
  if (showColumnHeaders) {
    const headerRow: string[] = [];
    for (let col = 0; col < numCols; col++) {
      const colLetter = String.fromCharCode(65 + startCol + col);
      headerRow.push(colLetter);
      columnWidths[col] = Math.max(columnWidths[col]!, colLetter.length);
    }
    cellValues.push(headerRow);
  }

  // Collect all cell data
  for (let row = startRow; row < startRow + numRows; row++) {
    const rowValues: string[] = [];
    for (let col = 0; col < numCols; col++) {
      let value: string;
      if (targetSheetName) {
        const cellValue = engine.getCellValue({ 
          sheetName: targetSheetName, 
          rowIndex: row - 1, 
          colIndex: startCol + col 
        });
        value = cellValue === "" || cellValue == null ? emptyCellChar : String(cellValue);
      } else {
        // No sheets available
        value = emptyCellChar;
      }
      
      rowValues.push(value);
      // Update column width, respecting max width
      columnWidths[col] = Math.max(
        columnWidths[col]!, 
        Math.min(value.length, maxColWidth)
      );
    }
    cellValues.push(rowValues);
  }

  // Calculate total row width for consistent formatting
  let totalRowWidth = 0;
  if (showRowNumbers) {
    totalRowWidth += 3; // Row number width
  }
  for (let col = 0; col < numCols; col++) {
    totalRowWidth += columnWidths[col]!;
  }
  // Add separator characters: " | " between columns (3 chars each, but last column doesn't have trailing separator)
  if (numCols > 0) {
    totalRowWidth += (numCols - 1) * 3; // " | " between columns
    if (showRowNumbers) {
      totalRowWidth += 3; // " | " after row number
    }
  }

  // Second pass: format the output with proper column widths
  let result = "";
  
  for (let i = 0; i < cellValues.length; i++) {
    const rowValues = cellValues[i]!;
    const formattedCells: string[] = [];
    
    // Add row number if requested
    if (showRowNumbers) {
      if (i === 0 && showColumnHeaders) {
        formattedCells.push("   "); // Empty space for header row
      } else {
        const rowNum = showColumnHeaders ? startRow + i - 1 : startRow + i;
        formattedCells.push(rowNum.toString().padStart(3));
      }
    }
    
    // Format each cell with proper width
    for (let col = 0; col < rowValues.length; col++) {
      let cellValue = rowValues[col]!;
      
      // Truncate if too long
      if (cellValue.length > maxColWidth) {
        cellValue = cellValue.substring(0, maxColWidth - 3) + "...";
      }
      
      // Pad to column width
      formattedCells.push(cellValue.padEnd(columnWidths[col]!));
    }
    
    // Join with consistent separators and ensure exact row width
    let rowContent = formattedCells.join(" | ");
    // Pad the entire row to ensure consistent width
    rowContent = rowContent.padEnd(totalRowWidth);
    result += rowContent + "\n";
    
    // Add separator line after header
    if (i === 0 && showColumnHeaders) {
      const separatorParts: string[] = [];
      
      if (showRowNumbers) {
        separatorParts.push("---");
      }
      
      for (let col = 0; col < numCols; col++) {
        separatorParts.push("-".repeat(columnWidths[col]!));
      }
      
      // Create separator line with same width as data rows
      let separatorLine = separatorParts.join("-+-");
      separatorLine = separatorLine.padEnd(totalRowWidth);
      result += separatorLine + "\n";
    }
  }
  
  return result;
}


