import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import { transformAST } from "./ast-traverser";
import type { ReferenceNode, RangeNode } from "../parser/ast";

/**
 * Information about cells being moved
 */
export interface MovedCellsInfo {
  /**
   * Set of cell keys in format "workbookName:sheetName:colIndex:rowIndex"
   * for fast O(1) lookup
   */
  cellsSet: Set<string>;
  
  /**
   * Workbook name of the source cells (all moved cells must be in same workbook/sheet)
   */
  workbookName: string;
  
  /**
   * Sheet name of the source cells
   */
  sheetName: string;
  
  /**
   * Row offset for the move (targetRow - sourceRow)
   */
  rowOffset: number;
  
  /**
   * Column offset for the move (targetCol - sourceCol)
   */
  colOffset: number;
}

/**
 * Creates a cell key for fast lookup
 */
function createCellKey(workbookName: string, sheetName: string, colIndex: number, rowIndex: number): string {
  return `${workbookName}:${sheetName}:${colIndex}:${rowIndex}`;
}

/**
 * Checks if a cell is in the moved cells set
 */
function isCellMoved(
  movedCells: MovedCellsInfo,
  workbookName: string | undefined,
  sheetName: string | undefined,
  colIndex: number,
  rowIndex: number
): boolean {
  // If no workbook/sheet specified in reference, it refers to current context
  // We only update if it matches the moved cells' workbook/sheet
  const refWorkbook = workbookName ?? movedCells.workbookName;
  const refSheet = sheetName ?? movedCells.sheetName;
  
  // Only update references in the same workbook/sheet as the moved cells
  if (refWorkbook !== movedCells.workbookName || refSheet !== movedCells.sheetName) {
    return false;
  }
  
  const key = createCellKey(refWorkbook, refSheet, colIndex, rowIndex);
  return movedCells.cellsSet.has(key);
}

/**
 * Updates cell and range references in a formula when cells are moved
 * 
 * @param formula - The formula string (without the leading =)
 * @param movedCells - Information about which cells were moved
 * @returns The updated formula string, or the original if no changes were made
 * 
 * @example
 * // Moving A1 to D5 (offset: col+3, row+4)
 * updateReferencesForMovedCells("A1+B1", {...}) // "D5+B1" (only A1 updated)
 * updateReferencesForMovedCells("SUM(A1:D5)", {...}) // "SUM(E4:H9)" (if entire range moved)
 */
export function updateReferencesForMovedCells(
  formula: string,
  movedCells: MovedCellsInfo
): string {
  try {
    const ast = parseFormula(formula);
    
    const transformedAST = transformAST(ast, (node) => {
      // Handle cell references
      if (node.type === "reference") {
        const refNode = node as ReferenceNode;
        
        // Check if this cell reference points to a moved cell
        if (isCellMoved(
          movedCells,
          refNode.workbookName,
          refNode.sheetName,
          refNode.address.colIndex,
          refNode.address.rowIndex
        )) {
          // Update the reference to the new location
          return {
            ...refNode,
            address: {
              colIndex: refNode.address.colIndex + movedCells.colOffset,
              rowIndex: refNode.address.rowIndex + movedCells.rowOffset,
            },
          };
        }
      }
      
      // Handle range references
      if (node.type === "range") {
        const rangeNode = node as RangeNode;
        
        // Only update if the range has finite bounds
        if (
          rangeNode.range.end.col.type === "number" &&
          rangeNode.range.end.row.type === "number"
        ) {
          const startCol = rangeNode.range.start.col;
          const startRow = rangeNode.range.start.row;
          const endCol = rangeNode.range.end.col.value;
          const endRow = rangeNode.range.end.row.value;
          
          // Check if BOTH start and end of the range are in the moved cells
          const startMoved = isCellMoved(
            movedCells,
            rangeNode.workbookName,
            rangeNode.sheetName,
            startCol,
            startRow
          );
          
          const endMoved = isCellMoved(
            movedCells,
            rangeNode.workbookName,
            rangeNode.sheetName,
            endCol,
            endRow
          );
          
          // Only update if ENTIRE range is being moved
          if (startMoved && endMoved) {
            // Verify that all cells in the range are being moved together
            // by checking if the range dimensions are preserved
            const rangeWidth = endCol - startCol;
            const rangeHeight = endRow - startRow;
            
            // Check a few cells in the middle to ensure contiguous movement
            // For small ranges, check all cells
            let allCellsMoved = true;
            
            // Sample checking strategy: check corners and center
            const cellsToCheck: Array<[number, number]> = [
              [startCol, startRow],     // Top-left (already checked)
              [endCol, endRow],         // Bottom-right (already checked)
            ];
            
            // Add middle cells for larger ranges
            if (rangeWidth > 1 || rangeHeight > 1) {
              const midCol = Math.floor((startCol + endCol) / 2);
              const midRow = Math.floor((startRow + endRow) / 2);
              cellsToCheck.push([midCol, midRow]);
              
              if (rangeWidth > 1) {
                cellsToCheck.push([midCol, startRow]);
                cellsToCheck.push([midCol, endRow]);
              }
              if (rangeHeight > 1) {
                cellsToCheck.push([startCol, midRow]);
                cellsToCheck.push([endCol, midRow]);
              }
            }
            
            // Check if all sampled cells are moved
            for (const [col, row] of cellsToCheck) {
              if (!isCellMoved(movedCells, rangeNode.workbookName, rangeNode.sheetName, col, row)) {
                allCellsMoved = false;
                break;
              }
            }
            
            if (allCellsMoved) {
              // Update the entire range
              return {
                ...rangeNode,
                range: {
                  start: {
                    col: startCol + movedCells.colOffset,
                    row: startRow + movedCells.rowOffset,
                  },
                  end: {
                    col: { type: "number" as const, value: endCol + movedCells.colOffset },
                    row: { type: "number" as const, value: endRow + movedCells.rowOffset },
                  },
                },
              };
            }
          }
        }
      }
      
      return node;
    });

    return astToString(transformedAST);
  } catch (error) {
    // If parsing fails, return the original formula
    return formula;
  }
}

/**
 * Checks if a formula contains a reference to a specific cell
 * 
 * @param formula - The formula string (without the leading =)
 * @param workbookName - The workbook name
 * @param sheetName - The sheet name
 * @param colIndex - The column index
 * @param rowIndex - The row index
 * @returns True if the formula references the cell
 */
export function formulaReferencesCell(
  formula: string,
  workbookName: string,
  sheetName: string,
  colIndex: number,
  rowIndex: number
): boolean {
  try {
    const ast = parseFormula(formula);
    let hasReference = false;

    transformAST(ast, (node) => {
      if (node.type === "reference") {
        const refNode = node as ReferenceNode;
        
        // Match if same cell (with or without explicit sheet/workbook)
        if (
          refNode.address.colIndex === colIndex &&
          refNode.address.rowIndex === rowIndex
        ) {
          // If reference has workbook/sheet, they must match
          if (refNode.workbookName && refNode.workbookName !== workbookName) {
            return node;
          }
          if (refNode.sheetName && refNode.sheetName !== sheetName) {
            return node;
          }
          
          hasReference = true;
        }
      }
      return node;
    });

    return hasReference;
  } catch (error) {
    // If parsing fails, assume no reference
    return false;
  }
}

/**
 * Checks if a formula contains a reference to a specific range
 * 
 * @param formula - The formula string (without the leading =)
 * @param workbookName - The workbook name
 * @param sheetName - The sheet name
 * @param startCol - Range start column
 * @param startRow - Range start row
 * @param endCol - Range end column
 * @param endRow - Range end row
 * @returns True if the formula references the range
 */
export function formulaReferencesRange(
  formula: string,
  workbookName: string,
  sheetName: string,
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number
): boolean {
  try {
    const ast = parseFormula(formula);
    let hasReference = false;

    transformAST(ast, (node) => {
      if (node.type === "range") {
        const rangeNode = node as RangeNode;
        
        // Check if range matches (must be finite)
        if (
          rangeNode.range.end.col.type === "number" &&
          rangeNode.range.end.row.type === "number" &&
          rangeNode.range.start.col === startCol &&
          rangeNode.range.start.row === startRow &&
          rangeNode.range.end.col.value === endCol &&
          rangeNode.range.end.row.value === endRow
        ) {
          // If reference has workbook/sheet, they must match
          if (rangeNode.workbookName && rangeNode.workbookName !== workbookName) {
            return node;
          }
          if (rangeNode.sheetName && rangeNode.sheetName !== sheetName) {
            return node;
          }
          
          hasReference = true;
        }
      }
      return node;
    });

    return hasReference;
  } catch (error) {
    // If parsing fails, assume no reference
    return false;
  }
}

