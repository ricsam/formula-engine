import type {
  CellAddress,
  RangeAddress,
  SerializedCellValue,
  Sheet,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";
import type { WorkbookManager } from "./workbook-manager";

export type LookupOrder = "row-major" | "col-major";

export type RangeEvalOrderEntryDict = {
  value: {
    type: "value";
    address: CellAddress;
  };
  empty_cell: {
    type: "empty_cell";
    address: CellAddress;
    candidates: CellAddress[];
  };
  empty_range: {
    type: "empty_range";
    address: RangeAddress;
    candidates: CellAddress[];
  };
};

export type RangeEvalOrderEntry =
  | RangeEvalOrderEntryDict["value"]
  | RangeEvalOrderEntryDict["empty_cell"]
  | RangeEvalOrderEntryDict["empty_range"];

/**
 * Build a deterministic, ordered list describing every cell inside a lookup range.
 * This function analyzes the range and classifies each cell/range as either:
 * - A value-like cell (occupied)
 * - An empty cell with 0-2 frontier candidates
 * - An empty range with 0-2 frontier candidates
 *
 * @param this - WorkbookManager instance
 * @param lookupOrder - "row-major" or "col-major" iteration order
 * @param lookupRange - The range to analyze
 * @returns Ordered array of entries describing the range
 */
export function buildRangeEvalOrder(
  this: WorkbookManager,
  lookupOrder: LookupOrder,
  lookupRange: RangeAddress
): RangeEvalOrderEntry[] {
  const sheet = this.getSheet(lookupRange);
  if (!sheet) {
    throw new Error("Sheet not found");
  }

  const result: RangeEvalOrderEntry[] = [];

  // Get range bounds
  const startRow = lookupRange.range.start.row;
  const startCol = lookupRange.range.start.col;

  // Build a map of occupied cells within the range
  const occupiedCells = new Map<string, CellAddress>();
  let maxRow = startRow;
  let maxCol = startCol;
  
  for (const cellAddr of this.iterateCellsInRange(lookupRange)) {
    occupiedCells.set(getCellReference(cellAddr), cellAddr);
    maxRow = Math.max(maxRow, cellAddr.rowIndex);
    maxCol = Math.max(maxCol, cellAddr.colIndex);
  }

  // For infinite ranges, also consider ALL occupied cells in the sheet
  // (including frontier candidates and their spills) to determine extent
  if (
    lookupRange.range.end.row.type === "infinity" ||
    lookupRange.range.end.col.type === "infinity"
  ) {
    // Get all cells in the sheet to find the full extent
    const allCells = this.iterateCellsInRange({
      workbookName: lookupRange.workbookName,
      sheetName: lookupRange.sheetName,
      range: {
        start: { row: 0, col: 0 },
        end: {
          row: { type: "infinity", sign: "positive" },
          col: { type: "infinity", sign: "positive" },
        },
      },
    });

    for (const cellAddr of allCells) {
      // Consider all occupied cells (formulas and values) to determine extent
      maxRow = Math.max(maxRow, cellAddr.rowIndex);
      maxCol = Math.max(maxCol, cellAddr.colIndex);
    }
  }

  // Determine the effective end bounds
  // For finite ranges, use the specified end
  // For infinite ranges, use the maximum cell found (including frontier candidates)
  const endRow =
    lookupRange.range.end.row.type === "number"
      ? lookupRange.range.end.row.value
      : maxRow;
  const endCol =
    lookupRange.range.end.col.type === "number"
      ? lookupRange.range.end.col.value
      : maxCol;

  if (lookupOrder === "row-major") {
    // Iterate row by row, left to right
    for (let row = startRow; row <= endRow; row++) {
      processRowMajorRow.call(
        this,
        row,
        startCol,
        endCol,
        lookupRange,
        occupiedCells,
        sheet,
        result
      );

      // If the lookup range is infinite in columns, emit a final infinite range for this row ONLY
      if (lookupRange.range.end.col.type === "infinity") {
        const nextCol = endCol + 1;
        const firstCellBeyondEnd: CellAddress = {
          rowIndex: row,
          colIndex: nextCol,
          sheetName: lookupRange.sheetName,
          workbookName: lookupRange.workbookName,
        };
        
        const candidates = findCandidatesForCell.call(
          this,
          firstCellBeyondEnd,
          lookupRange,
          sheet,
          lookupOrder
        );

        result.push({
          type: "empty_range",
          address: {
            workbookName: lookupRange.workbookName,
            sheetName: lookupRange.sheetName,
            range: {
              start: {
                row: row,
                col: nextCol,
              },
              end: {
                // Always constrain to this row only in row-major
                row: { type: "number", value: row },
                col: lookupRange.range.end.col,
              },
            },
          },
          candidates,
        });
      }
    }

    // If the lookup range is infinite in rows, emit a final infinite range
    if (lookupRange.range.end.row.type === "infinity") {
      // Find candidates for cells beyond the last processed row
      const nextRow = endRow + 1;
      const firstCellInNextRow: CellAddress = {
        rowIndex: nextRow,
        colIndex: startCol,
        sheetName: lookupRange.sheetName,
        workbookName: lookupRange.workbookName,
      };
      
      const candidates = findCandidatesForCell.call(
        this,
        firstCellInNextRow,
        lookupRange,
        sheet,
        lookupOrder
      );

      result.push({
        type: "empty_range",
        address: {
          workbookName: lookupRange.workbookName,
          sheetName: lookupRange.sheetName,
          range: {
            start: {
              row: nextRow,
              col: startCol,
            },
            end: {
              row: lookupRange.range.end.row,
              col: lookupRange.range.end.col,
            },
          },
        },
        candidates,
      });
    }
  } else {
    // col-major: iterate column by column, top to bottom
    
    // Check if we can emit a single infinite range for the entire remaining region
    // This happens when both rows and columns are infinite and all cells have the same candidates
    if (
      lookupRange.range.end.row.type === "infinity" &&
      lookupRange.range.end.col.type === "infinity"
    ) {
      // Check if the entire range has uniform candidates
      const firstCell: CellAddress = {
        rowIndex: startRow,
        colIndex: startCol,
        sheetName: lookupRange.sheetName,
        workbookName: lookupRange.workbookName,
      };
      const candidates = findCandidatesForCell.call(
        this,
        firstCell,
        lookupRange,
        sheet,
        lookupOrder
      );

      // Check if all processed cells and the region beyond have the same candidates
      let hasOccupiedCells = false;
      for (let col = startCol; col <= endCol; col++) {
        for (let row = startRow; row <= endRow; row++) {
          const cellRef = getCellReference({
            rowIndex: row,
            colIndex: col,
            sheetName: lookupRange.sheetName,
            workbookName: lookupRange.workbookName,
          });
          if (occupiedCells.has(cellRef)) {
            hasOccupiedCells = true;
            break;
          }
        }
        if (hasOccupiedCells) break;
      }

      // If no occupied cells in the range, emit a single infinite range
      if (!hasOccupiedCells) {
        result.push({
          type: "empty_range",
          address: {
            workbookName: lookupRange.workbookName,
            sheetName: lookupRange.sheetName,
            range: {
              start: {
                row: startRow,
                col: startCol,
              },
              end: {
                row: lookupRange.range.end.row,
                col: lookupRange.range.end.col,
              },
            },
          },
          candidates,
        });
        return result;
      }
    }

    for (let col = startCol; col <= endCol; col++) {
      processColMajorColumn.call(
        this,
        col,
        startRow,
        endRow,
        lookupRange,
        occupiedCells,
        sheet,
        result
      );

      // If the lookup range is infinite in rows, emit a final infinite range for this column ONLY
      if (lookupRange.range.end.row.type === "infinity") {
        const nextRow = endRow + 1;
        const firstCellBeyondEnd: CellAddress = {
          rowIndex: nextRow,
          colIndex: col,
          sheetName: lookupRange.sheetName,
          workbookName: lookupRange.workbookName,
        };
        
        const candidates = findCandidatesForCell.call(
          this,
          firstCellBeyondEnd,
          lookupRange,
          sheet,
          lookupOrder
        );

        result.push({
          type: "empty_range",
          address: {
            workbookName: lookupRange.workbookName,
            sheetName: lookupRange.sheetName,
            range: {
              start: {
                row: nextRow,
                col: col,
              },
              end: {
                row: lookupRange.range.end.row,
                // Always constrain to this column only in col-major
                col: { type: "number", value: col },
              },
            },
          },
          candidates,
        });
      }
    }

    // If the lookup range is infinite in columns, emit a final infinite range
    if (lookupRange.range.end.col.type === "infinity") {
      // Find candidates for cells beyond the last processed column
      const nextCol = endCol + 1;
      const firstCellInNextCol: CellAddress = {
        rowIndex: startRow,
        colIndex: nextCol,
        sheetName: lookupRange.sheetName,
        workbookName: lookupRange.workbookName,
      };
      
      const candidates = findCandidatesForCell.call(
        this,
        firstCellInNextCol,
        lookupRange,
        sheet,
        lookupOrder
      );

      result.push({
        type: "empty_range",
        address: {
          workbookName: lookupRange.workbookName,
          sheetName: lookupRange.sheetName,
          range: {
            start: {
              row: startRow,
              col: nextCol,
            },
            end: {
              row: lookupRange.range.end.row,
              col: lookupRange.range.end.col,
            },
          },
        },
        candidates,
      });
    }
  }

  return result;
}

/**
 * Process a single row in row-major order
 */
function processRowMajorRow(
  this: WorkbookManager,
  row: number,
  startCol: number,
  endCol: number,
  lookupRange: RangeAddress,
  occupiedCells: Map<string, CellAddress>,
  sheet: Sheet,
  result: RangeEvalOrderEntry[]
): void {
  let col = startCol;

  while (col <= endCol) {
    const cellAddr: CellAddress = {
      rowIndex: row,
      colIndex: col,
      sheetName: lookupRange.sheetName,
      workbookName: lookupRange.workbookName,
    };
    const cellRef = getCellReference(cellAddr);

    if (occupiedCells.has(cellRef)) {
      // This cell is occupied (has a value or formula)
      result.push({
        type: "value",
        address: cellAddr,
      });
      col++;
    } else {
      // This cell is empty - find the contiguous empty range in this row with same candidates
      const emptyRangeStart = col;
      let emptyRangeEnd = col;

      // Get candidates for the first cell
      const firstCandidates = findCandidatesForCell.call(
        this,
        cellAddr,
        lookupRange,
        sheet,
        "row-major"
      );

      // Try to extend the range while candidates remain the same
      while (emptyRangeEnd + 1 <= endCol) {
        const nextCellAddr: CellAddress = {
          rowIndex: row,
          colIndex: emptyRangeEnd + 1,
          sheetName: lookupRange.sheetName,
          workbookName: lookupRange.workbookName,
        };
        const nextCellRef = getCellReference(nextCellAddr);

        // Stop if next cell is occupied
        if (occupiedCells.has(nextCellRef)) {
          break;
        }

        // Get candidates for next cell
        const nextCandidates = findCandidatesForCell.call(
          this,
          nextCellAddr,
          lookupRange,
          sheet,
          "row-major"
        );

        // Stop if candidates differ
        if (!candidatesEqual(firstCandidates, nextCandidates)) {
          break;
        }

        emptyRangeEnd++;
      }

      // If the empty range is a single cell, emit as empty_cell
      if (emptyRangeStart === emptyRangeEnd) {
        result.push({
          type: "empty_cell",
          address: cellAddr,
          candidates: firstCandidates,
        });
      } else {
        // Emit as empty_range
        result.push({
          type: "empty_range",
          address: {
            workbookName: lookupRange.workbookName,
            sheetName: lookupRange.sheetName,
            range: {
              start: {
                row,
                col: emptyRangeStart,
              },
              end: {
                row: { type: "number", value: row },
                col: { type: "number", value: emptyRangeEnd },
              },
            },
          },
          candidates: firstCandidates,
        });
      }

      col = emptyRangeEnd + 1;
    }
  }
}

/**
 * Process a single column in col-major order
 */
function processColMajorColumn(
  this: WorkbookManager,
  col: number,
  startRow: number,
  endRow: number,
  lookupRange: RangeAddress,
  occupiedCells: Map<string, CellAddress>,
  sheet: Sheet,
  result: RangeEvalOrderEntry[]
): void {
  let row = startRow;

  while (row <= endRow) {
    const cellAddr: CellAddress = {
      rowIndex: row,
      colIndex: col,
      sheetName: lookupRange.sheetName,
      workbookName: lookupRange.workbookName,
    };
    const cellRef = getCellReference(cellAddr);

    if (occupiedCells.has(cellRef)) {
      // This cell is occupied (has a value or formula)
      result.push({
        type: "value",
        address: cellAddr,
      });
      row++;
    } else {
      // This cell is empty - find the contiguous empty range in this column with same candidates
      const emptyRangeStart = row;
      let emptyRangeEnd = row;

      // Get candidates for the first cell
      const firstCandidates = findCandidatesForCell.call(
        this,
        cellAddr,
        lookupRange,
        sheet,
        "col-major"
      );

      // Try to extend the range while candidates remain the same
      while (emptyRangeEnd + 1 <= endRow) {
        const nextCellAddr: CellAddress = {
          rowIndex: emptyRangeEnd + 1,
          colIndex: col,
          sheetName: lookupRange.sheetName,
          workbookName: lookupRange.workbookName,
        };
        const nextCellRef = getCellReference(nextCellAddr);

        // Stop if next cell is occupied
        if (occupiedCells.has(nextCellRef)) {
          break;
        }

        // Get candidates for next cell
        const nextCandidates = findCandidatesForCell.call(
          this,
          nextCellAddr,
          lookupRange,
          sheet,
          "col-major"
        );

        // Stop if candidates differ
        if (!candidatesEqual(firstCandidates, nextCandidates)) {
          break;
        }

        emptyRangeEnd++;
      }

      // If the empty range is a single cell, emit as empty_cell
      if (emptyRangeStart === emptyRangeEnd) {
        result.push({
          type: "empty_cell",
          address: cellAddr,
          candidates: firstCandidates,
        });
      } else {
        // Emit as empty_range
        result.push({
          type: "empty_range",
          address: {
            workbookName: lookupRange.workbookName,
            sheetName: lookupRange.sheetName,
            range: {
              start: {
                row: emptyRangeStart,
                col,
              },
              end: {
                row: { type: "number", value: emptyRangeEnd },
                col: { type: "number", value: col },
              },
            },
          },
          candidates: firstCandidates,
        });
      }

      row = emptyRangeEnd + 1;
    }
  }
}

/**
 * Check if two candidate arrays are equal
 */
function candidatesEqual(a: CellAddress[], b: CellAddress[]): boolean {
  if (a.length !== b.length) return false;

  for (let i = 0; i < a.length; i++) {
    const addrA = a[i];
    const addrB = b[i];
    if (!addrA || !addrB) return false;
    if (
      addrA.rowIndex !== addrB.rowIndex ||
      addrA.colIndex !== addrB.colIndex ||
      addrA.sheetName !== addrB.sheetName ||
      addrA.workbookName !== addrB.workbookName
    ) {
      return false;
    }
  }

  return true;
}

/**
 * Find frontier candidates for a single empty cell
 * Returns candidates that could spill into this cell:
 * - If there's a direct left or above candidate, return those (max 2)
 * - Otherwise, return ALL step-pattern diagonal candidates
 */
function findCandidatesForCell(
  this: WorkbookManager,
  cellAddr: CellAddress,
  lookupRange: RangeAddress,
  sheet: Sheet,
  lookupOrder: LookupOrder
): CellAddress[] {
  // Find nearest-left anchor (same row)
  const leftCandidate = findNearestLeftAnchor.call(
    this,
    cellAddr,
    lookupRange,
    sheet
  );

  // Find nearest-above anchor (same column)
  const aboveCandidate = findNearestAboveAnchor.call(
    this,
    cellAddr,
    lookupRange,
    sheet
  );

  // If we have direct candidates (left and/or above), use those
  if (leftCandidate || aboveCandidate) {
    const candidates: CellAddress[] = [];
    if (lookupOrder === "row-major") {
      if (leftCandidate) candidates.push(leftCandidate);
      if (aboveCandidate) candidates.push(aboveCandidate);
    } else {
      if (aboveCandidate) candidates.push(aboveCandidate);
      if (leftCandidate) candidates.push(leftCandidate);
    }
    return candidates;
  }

  // No direct candidates - find ALL step-pattern diagonal candidates
  const diagonalCandidates = findAllDiagonalStepCandidates.call(
    this,
    cellAddr,
    lookupRange,
    sheet
  );

  return diagonalCandidates;
}

/**
 * Check if two cell addresses refer to the same cell
 */
function isSameCell(a: CellAddress, b: CellAddress): boolean {
  return (
    a.rowIndex === b.rowIndex &&
    a.colIndex === b.colIndex &&
    a.sheetName === b.sheetName &&
    a.workbookName === b.workbookName
  );
}

/**
 * Find the nearest anchor (formula cell) to the left of the target cell
 * within the same row
 * Only finds cells with formulas (starting with "=")
 * Searches both inside and outside the lookup range boundaries
 */
function findNearestLeftAnchor(
  this: WorkbookManager,
  targetCell: CellAddress,
  lookupRange: RangeAddress,
  sheet: Sheet
): CellAddress | null {
  const row = targetCell.rowIndex;
  const targetCol = targetCell.colIndex;

  // Search from targetCol-1 down to column 0 (search entire row to the left)
  for (let col = targetCol - 1; col >= 0; col--) {
    const cellAddr: CellAddress = {
      rowIndex: row,
      colIndex: col,
      sheetName: targetCell.sheetName,
      workbookName: targetCell.workbookName,
    };
    const cellRef = getCellReference(cellAddr);
    const content = sheet.content.get(cellRef);

    // Only consider formula cells (starting with "=")
    if (typeof content === "string" && content.startsWith("=")) {
      return cellAddr;
    }
  }

  return null;
}

/**
 * Find the nearest anchor (formula cell) above the target cell
 * within the same column
 * Only finds cells with formulas (starting with "=")
 * Searches outside the lookup range boundaries
 */
function findNearestAboveAnchor(
  this: WorkbookManager,
  targetCell: CellAddress,
  lookupRange: RangeAddress,
  sheet: Sheet
): CellAddress | null {
  const col = targetCell.colIndex;
  const targetRow = targetCell.rowIndex;

  // Search from targetRow-1 up to row 0 (search entire column above)
  for (let row = targetRow - 1; row >= 0; row--) {
    const cellAddr: CellAddress = {
      rowIndex: row,
      colIndex: col,
      sheetName: targetCell.sheetName,
      workbookName: targetCell.workbookName,
    };
    const cellRef = getCellReference(cellAddr);
    const content = sheet.content.get(cellRef);

    // Only consider formula cells (starting with "=")
    if (typeof content === "string" && content.startsWith("=")) {
      return cellAddr;
    }
  }

  return null;
}

/**
 * Find ALL step-pattern candidates in the top-left diagonal quadrant
 * These are formulas that could spill diagonally to reach the target cell
 * and don't block each other (forming a "staircase" pattern)
 * Searches outside the lookup range boundaries (like direct left/above search)
 * 
 * For example, for target D7, candidates might be A6, B5, C4
 * - A6 at (5,0) could spill right and down to reach D7
 * - B5 at (4,1) could spill right and down to reach D7
 * - C4 at (3,2) could spill right and down to reach D7
 * - None blocks the others because each is further from origin (A1)
 */
function findAllDiagonalStepCandidates(
  this: WorkbookManager,
  targetCell: CellAddress,
  lookupRange: RangeAddress,
  sheet: Sheet
): CellAddress[] {
  const targetRow = targetCell.rowIndex;
  const targetCol = targetCell.colIndex;

  // Find all formula cells in the top-left quadrant (search entire grid to the top-left)
  const allCandidates: CellAddress[] = [];
  
  for (let row = targetRow - 1; row >= 0; row--) {
    for (let col = targetCol - 1; col >= 0; col--) {
      const cellAddr: CellAddress = {
        rowIndex: row,
        colIndex: col,
        sheetName: targetCell.sheetName,
        workbookName: targetCell.workbookName,
      };
      const cellRef = getCellReference(cellAddr);
      const content = sheet.content.get(cellRef);

      // Only consider formula cells (starting with "=")
      if (typeof content === "string" && content.startsWith("=")) {
        allCandidates.push(cellAddr);
      }
    }
  }

  // Filter out candidates that are blocked by other candidates
  // A candidate at (r1, c1) is blocked by another at (r2, c2) if:
  // - (r2, c2) is between (r1, c1) and the target
  // - Meaning: r1 <= r2 < targetRow and c1 <= c2 < targetCol
  // - Since spills go down-right, a formula can only block cells at positions >= itself
  const unblockedCandidates = allCandidates.filter((candidate) => {
    // Check if this candidate is blocked by any other candidate
    for (const other of allCandidates) {
      if (isSameCell(candidate, other)) continue;
      
      // Check if 'other' blocks 'candidate' from reaching target
      // 'other' blocks 'candidate' if 'other' is between 'candidate' and target
      // For rectangular spills going down-right:
      // 'other' at (r2, c2) blocks 'candidate' at (r1, c1) if:
      // r1 <= r2 AND c1 <= c2 (other is down-right of candidate, potentially in its spill path)
      // AND the spill from other would reach the candidate's row or column before target
      
      // Actually, simpler: 'other' blocks 'candidate' if 'other' would occupy
      // a cell that 'candidate' needs to spill through to reach target
      // Since we're looking for non-blocking step patterns, candidates should satisfy:
      // For any two candidates (r1,c1) and (r2,c2), neither should be in the 
      // minimal spill rectangle of the other that reaches the target
      
      // A simpler approach: candidates form valid steps if for any two candidates,
      // one is strictly top-left of the other (both row and col smaller)
      // This ensures they don't block each other
      
      // Check if 'other' could block 'candidate'
      if (
        other.rowIndex >= candidate.rowIndex &&
        other.colIndex >= candidate.colIndex &&
        other.rowIndex < targetRow &&
        other.colIndex < targetCol
      ) {
        // 'other' is in the potential spill path from 'candidate' to target
        return false; // candidate is blocked
      }
    }
    return true; // candidate is not blocked
  });

  return unblockedCandidates;
}
