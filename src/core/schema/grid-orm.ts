/**
 * GridOrm - Object-Relational Mapping for grid/range data
 *
 * Provides read/write operations for a 2D range of cells with type-safe access.
 * Supports both column-major and row-major array access patterns.
 */

import type {
  CellAddress,
  FiniteSpreadsheetRange,
  SerializedCellValue,
} from "../types";
import type { FormulaEngine } from "../engine";

export type GridParseFunction<TCellMetadata, TValue> = (
  value: unknown,
  metadata: TCellMetadata
) => TValue;

export type GridWriteFunction<TCellMetadata, TValue> = (value: TValue) => {
  value: SerializedCellValue;
  metadata?: TCellMetadata;
};

export class GridOrm<TValue, TCellMetadata = unknown> {
  private numCols: number;
  private numRows: number;

  constructor(
    private engine: FormulaEngine<any, any>,
    private workbookName: string,
    private sheetName: string,
    private range: FiniteSpreadsheetRange,
    private parser: GridParseFunction<TCellMetadata, TValue>,
    private writer: GridWriteFunction<TCellMetadata, TValue>,
    private namespace: string
  ) {
    this.numCols = range.end.col - range.start.col + 1;
    this.numRows = range.end.row - range.start.row + 1;
  }

  /**
   * Get the cell address for a given column and row index within the grid
   */
  private getCellAddress(colIndex: number, rowIndex: number): CellAddress {
    return {
      workbookName: this.workbookName,
      sheetName: this.sheetName,
      colIndex: this.range.start.col + colIndex,
      rowIndex: this.range.start.row + rowIndex,
    };
  }

  /**
   * Read a single cell and parse it
   */
  private readCell(colIndex: number, rowIndex: number): TValue {
    const cellAddress = this.getCellAddress(colIndex, rowIndex);
    const value = this.engine.getCellValue(cellAddress);
    const metadata = this.engine.getCellMetadata(cellAddress) as TCellMetadata;
    return this.parser(value, metadata);
  }

  /**
   * Write a single cell using the write function
   */
  private writeCell(colIndex: number, rowIndex: number, value: TValue): void {
    const cellAddress = this.getCellAddress(colIndex, rowIndex);
    const { value: serializedValue, metadata } = this.writer(value);
    this.engine.setCellContent(cellAddress, serializedValue);
    if (metadata !== undefined) {
      this.engine.setCellMetadata(cellAddress, metadata);
    }
  }

  /**
   * Get all cells as a column-major 2D array (readonly)
   * columns[colIndex][rowIndex]
   */
  get columns(): readonly (readonly TValue[])[] {
    const result: TValue[][] = [];
    for (let col = 0; col < this.numCols; col++) {
      const column: TValue[] = [];
      for (let row = 0; row < this.numRows; row++) {
        column.push(this.readCell(col, row));
      }
      result.push(column);
    }
    return result;
  }

  /**
   * Get all cells as a row-major 2D array (readonly)
   * rows[rowIndex][colIndex]
   */
  get rows(): readonly (readonly TValue[])[] {
    const result: TValue[][] = [];
    for (let row = 0; row < this.numRows; row++) {
      const rowData: TValue[] = [];
      for (let col = 0; col < this.numCols; col++) {
        rowData.push(this.readCell(col, row));
      }
      result.push(rowData);
    }
    return result;
  }

  /**
   * Set a single value at the specified position
   * Position is relative to the grid (0-indexed)
   */
  setValue(value: TValue, position: { col: number; row: number }): void {
    if (
      position.col < 0 ||
      position.col >= this.numCols ||
      position.row < 0 ||
      position.row >= this.numRows
    ) {
      throw new Error(
        `Grid schema "${this.namespace}": Position (${position.col}, ${position.row}) is out of bounds. Grid is ${this.numCols}x${this.numRows}.`
      );
    }
    this.writeCell(position.col, position.row, value);
  }

  /**
   * Get a single value at the specified position
   * Position is relative to the grid (0-indexed)
   */
  getValue(position: { col: number; row: number }): TValue {
    if (
      position.col < 0 ||
      position.col >= this.numCols ||
      position.row < 0 ||
      position.row >= this.numRows
    ) {
      throw new Error(
        `Grid schema "${this.namespace}": Position (${position.col}, ${position.row}) is out of bounds. Grid is ${this.numCols}x${this.numRows}.`
      );
    }
    return this.readCell(position.col, position.row);
  }
}
