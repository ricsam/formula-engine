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

export class CellOrm<TValue, TCellMetadata = unknown> {
  constructor(
    private engine: FormulaEngine<any, any>,
    private cellAddress: CellAddress,
    private parser: CellParseFunction<TCellMetadata, TValue>,
    private namespace: string
  ) {}

  /**
   * Read the cell value and parse it using the schema's parse function
   */
  read(): TValue {
    const value = this.engine.getCellValue(this.cellAddress);
    const metadata = this.engine.getCellMetadata(this.cellAddress) as TCellMetadata;
    return this.parser(value, metadata);
  }

  /**
   * Write a value to the cell
   * The value should be the parsed type - it will be converted to a serializable value
   */
  write(value: TValue): void {
    // Convert TValue to SerializedCellValue
    let serializedValue: SerializedCellValue;

    if (
      typeof value === "string" ||
      typeof value === "number" ||
      typeof value === "boolean"
    ) {
      serializedValue = value;
    } else if (value === null || value === undefined) {
      serializedValue = undefined;
    } else if (typeof value === "object") {
      // For objects, we need to extract the actual cell value
      // This assumes the parsed type might be an object with a 'value' property
      const obj = value as Record<string, unknown>;
      if ("value" in obj) {
        const innerValue = obj.value;
        if (
          typeof innerValue === "string" ||
          typeof innerValue === "number" ||
          typeof innerValue === "boolean"
        ) {
          serializedValue = innerValue;
        } else {
          serializedValue = String(innerValue);
        }
      } else {
        // Fallback: convert to string
        serializedValue = JSON.stringify(value);
      }
    } else {
      serializedValue = String(value);
    }

    this.engine.setCellContent(this.cellAddress, serializedValue);
  }

  /**
   * Get the cell address
   */
  getAddress(): CellAddress {
    return { ...this.cellAddress };
  }
}
