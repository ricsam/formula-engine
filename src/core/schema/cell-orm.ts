/**
 * CellOrm - Object-Relational Mapping for single cell data
 *
 * Provides read/write operations for a single cell with type-safe access.
 * Used as the `this` context for custom cell API methods.
 */

import type { CellAddress, SerializedCellValue } from "../types";
import type { FormulaEngine } from "../engine";

export type CellParseFunction<TCellMetadata, TValue> = (
  value: unknown,
  metadata: TCellMetadata
) => TValue;

export type CellWriteFunction<TCellMetadata, TValue> = (value: TValue) => {
  value: SerializedCellValue;
  metadata?: TCellMetadata;
};

export class CellOrm<TValue, TCellMetadata = unknown> {
  constructor(
    private engine: FormulaEngine<any, any>,
    private cellAddress: CellAddress,
    private parser: CellParseFunction<TCellMetadata, TValue>,
    private writer: CellWriteFunction<TCellMetadata, TValue>,
    private namespace: string
  ) {}

  /**
   * Read the cell value and parse it using the schema's parse function
   */
  read(): TValue {
    const value = this.engine.getCellValue(this.cellAddress);
    const metadata = this.engine.getCellMetadata(
      this.cellAddress
    ) as TCellMetadata;
    return this.parser(value, metadata);
  }

  /**
   * Write a value to the cell using the schema's write function
   */
  write(value: TValue): void {
    const { value: serializedValue, metadata } = this.writer(value);
    this.engine.setCellContent(this.cellAddress, serializedValue);
    if (metadata !== undefined) {
      this.engine.setCellMetadata(this.cellAddress, metadata);
    }
  }

  /**
   * Get the cell address
   */
  getAddress(): CellAddress {
    return { ...this.cellAddress };
  }
}
