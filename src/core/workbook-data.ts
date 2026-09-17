/**
 * Serializable description of a whole workbook.
 *
 * `FormulaEngine.addWorkbook({ workbookName, data })` consumes this shape, which
 * lets an importer describe sheets, tables, names, styles and metadata without
 * driving the mutation APIs one call at a time. Areas are sheet-scoped rather
 * than workbook-scoped: the workbook name is supplied at import time, so the
 * same data can be added under any name.
 */

import type {
  CellDataType,
  CellStyle,
  RangeAddress,
  SerializedCellValue,
  SpreadsheetRange,
  SpreadsheetRangeEnd,
  StyleCondition,
} from "./types";

/**
 * A range within one sheet of the workbook being imported.
 */
export interface WorkbookDataArea {
  sheetName: string;
  range: SpreadsheetRange;
}

/**
 * Cell content and metadata for a single sheet.
 *
 * Sheets are created in array order, so the array order becomes the sheet tab
 * order.
 */
export interface WorkbookDataSheet<
  TCellMetadata = unknown,
  TSheetMetadata = unknown
> {
  name: string;
  /**
   * Cell content keyed by A1 reference. Values beginning with `=` are stored as
   * formulas. Plain objects are accepted for JSON round-tripping.
   */
  content?:
    | Map<string, SerializedCellValue>
    | Record<string, SerializedCellValue>;
  /** Per-cell consumer metadata keyed by A1 reference. */
  cellMetadata?: Map<string, TCellMetadata> | Record<string, TCellMetadata>;
  /** Sheet-level consumer metadata. */
  sheetMetadata?: TSheetMetadata;
}

/**
 * A table definition. Headers are read from the sheet content at `start`, so
 * the owning sheet's content is applied before tables are created.
 */
export interface WorkbookDataTable {
  name: string;
  sheetName: string;
  /** A1 reference of the header row's first cell, for example `"A1"`. */
  start: string;
  /** Number of data rows below the header row. */
  numRows: SpreadsheetRangeEnd;
  numCols: number;
}

/**
 * A named expression. Omitting `sheetName` scopes the name to the workbook.
 */
export interface WorkbookDataNamedExpression {
  name: string;
  /** Formula body without a leading `=`, for example `"0.25"` or `"Sheet1!$A$1"`. */
  expression: string;
  sheetName?: string;
}

export interface WorkbookDataCellStyle {
  areas: WorkbookDataArea[];
  style: CellStyle;
}

export interface WorkbookDataConditionalStyle {
  areas: WorkbookDataArea[];
  condition: StyleCondition;
}

export interface WorkbookDataCellDataType {
  areas: WorkbookDataArea[];
  dataType: CellDataType;
}

export interface WorkbookDataRangeMetadata<TRangeMetadata = unknown> {
  id?: string;
  areas: WorkbookDataArea[];
  metadata: TRangeMetadata;
}

export interface WorkbookData<
  TCellMetadata = unknown,
  TSheetMetadata = unknown,
  TWorkbookMetadata = unknown,
  TRangeMetadata = unknown
> {
  sheets: WorkbookDataSheet<TCellMetadata, TSheetMetadata>[];
  tables?: WorkbookDataTable[];
  namedExpressions?: WorkbookDataNamedExpression[];
  cellStyles?: WorkbookDataCellStyle[];
  conditionalStyles?: WorkbookDataConditionalStyle[];
  cellDataTypes?: WorkbookDataCellDataType[];
  rangeMetadata?: WorkbookDataRangeMetadata<TRangeMetadata>[];
  workbookMetadata?: TWorkbookMetadata;
}

/**
 * Accept either a Map or a plain object so `WorkbookData` survives a JSON round
 * trip.
 */
export function toEntryMap<TValue>(
  source: Map<string, TValue> | Record<string, TValue> | undefined
): Map<string, TValue> {
  if (!source) {
    return new Map();
  }
  if (source instanceof Map) {
    return source;
  }
  return new Map(Object.entries(source));
}

export function areaToRangeAddress(
  area: WorkbookDataArea,
  workbookName: string
): RangeAddress {
  return {
    workbookName,
    sheetName: area.sheetName,
    range: area.range,
  };
}

export function areasToRangeAddresses(
  areas: WorkbookDataArea[],
  workbookName: string
): RangeAddress[] {
  return areas.map((area) => areaToRangeAddress(area, workbookName));
}
