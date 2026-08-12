/**
 * StyleManager - Manages conditional styling for cells
 */

import type {
  CellAddress,
  CellDataType,
  CellStyle,
  ConditionalStyle,
  DirectCellDataType,
  DirectCellStyle,
  RangeAddress,
  SerializedCellValue,
} from "../types";
import type { StyleManagerSnapshot } from "../engine-snapshot";
import type { WorkbookManager } from "./workbook-manager";
import type { EvaluationManager } from "./evaluation-manager";
import { isCellInRange } from "../utils";
import {
  calculateGradientFactor,
  interpolateLCH,
  lchToHex,
} from "../utils/color-utils";
import {
  subtractRange,
  rangesIntersect,
  isRangeContained,
  intersectRanges,
} from "../utils/range-utils";
import {
  MutationObserverDispatcher,
  applyIndexedChanges,
  type IndexedMutationValue,
  type MutationDirection,
} from "./mutation-observer";

export type ConditionalStyleDataChange = {
  readonly kind: "conditional-style";
  readonly before?: IndexedMutationValue<ConditionalStyle>;
  readonly after?: IndexedMutationValue<ConditionalStyle>;
};

export type CellStyleDataChange = {
  readonly kind: "cell-style";
  readonly before?: IndexedMutationValue<DirectCellStyle>;
  readonly after?: IndexedMutationValue<DirectCellStyle>;
};

export type CellDataTypeDataChange = {
  readonly kind: "cell-data-type";
  readonly before?: IndexedMutationValue<DirectCellDataType>;
  readonly after?: IndexedMutationValue<DirectCellDataType>;
};

export type StyleDataChange =
  | ConditionalStyleDataChange
  | CellStyleDataChange
  | CellDataTypeDataChange;

export type StyleMutationObserver = (
  changes: readonly StyleDataChange[]
) => void;

type StyleCollectionKind = StyleDataChange["kind"];

const normalizeCellStyle = (style: CellStyle): CellStyle =>
  Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined)
  ) as CellStyle;

const normalizeDirectCellStyle = (style: DirectCellStyle): DirectCellStyle => ({
  ...style,
  style: normalizeCellStyle(style.style),
});

export class StyleManager {
  private conditionalStyles: ConditionalStyle[] = [];
  private cellStyles: DirectCellStyle[] = [];
  private cellDataTypes: DirectCellDataType[] = [];
  private readonly mutationDispatcher: MutationObserverDispatcher<StyleDataChange>;
  private mutationBatchDepth = 0;
  private mutationBatchBefore?: {
    conditionalStyles: ConditionalStyle[];
    cellStyles: DirectCellStyle[];
    cellDataTypes: DirectCellDataType[];
  };

  constructor(
    private evaluationManager: EvaluationManager,
    mutationObserver?: StyleMutationObserver,
    shouldObserve?: () => boolean,
    detachMutationValues = true,
    private readonly shouldBatchMutations: () => boolean = () => true
  ) {
    this.mutationDispatcher = new MutationObserverDispatcher(
      mutationObserver,
      shouldObserve,
      () => this.captureMutationBatchBefore(),
      detachMutationValues
    );
  }

  batchMutations<TResult>(callback: () => TResult): TResult {
    if (!this.shouldBatchMutations()) {
      return callback();
    }

    if (!this.mutationDispatcher.observed || this.mutationBatchDepth > 0) {
      this.mutationBatchDepth++;
      try {
        return callback();
      } finally {
        this.mutationBatchDepth--;
      }
    }

    this.mutationBatchBefore = undefined;
    this.mutationBatchDepth++;
    try {
      return this.mutationDispatcher.suppress(callback);
    } finally {
      this.mutationBatchDepth--;
      const before = this.takeMutationBatchBefore();
      if (before) {
        this.mutationDispatcher.report([
          ...this.diffCollection(
            "conditional-style",
            before.conditionalStyles,
            this.conditionalStyles
          ),
          ...this.diffCollection(
            "cell-style",
            before.cellStyles,
            this.cellStyles
          ),
          ...this.diffCollection(
            "cell-data-type",
            before.cellDataTypes,
            this.cellDataTypes
          ),
        ]);
      }
    }
  }

  private captureMutationBatchBefore(): void {
    if (this.mutationBatchDepth === 0 || this.mutationBatchBefore) {
      return;
    }
    this.mutationBatchBefore = {
      conditionalStyles: [...this.conditionalStyles],
      cellStyles: [...this.cellStyles],
      cellDataTypes: [...this.cellDataTypes],
    };
  }

  private takeMutationBatchBefore():
    | {
        conditionalStyles: ConditionalStyle[];
        cellStyles: DirectCellStyle[];
        cellDataTypes: DirectCellDataType[];
      }
    | undefined {
    const before = this.mutationBatchBefore;
    this.mutationBatchBefore = undefined;
    return before;
  }

  /**
   * Add a conditional style rule
   */
  addConditionalStyle(style: ConditionalStyle): void {
    const index = this.conditionalStyles.length;
    const after = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(style)
      : undefined;
    this.conditionalStyles.push(style);
    if (after) {
      this.mutationDispatcher.report([
        { kind: "conditional-style", after: { index, value: after } },
      ]);
    }
  }

  /**
   * Remove a conditional style rule by index for a specific workbook
   */
  removeConditionalStyle(workbookName: string, index: number): boolean {
    const workbookStyles = this.conditionalStyles.filter((style) =>
      style.areas.some((area) => area.workbookName === workbookName)
    );
    if (index < 0 || index >= workbookStyles.length) {
      return false;
    }
    // Find the actual index in the full array
    let currentIndex = 0;
    for (let i = 0; i < this.conditionalStyles.length; i++) {
      const style = this.conditionalStyles[i];
      if (
        style &&
        style.areas.some((area) => area.workbookName === workbookName)
      ) {
        if (currentIndex === index) {
          const before = this.mutationDispatcher.observed
            ? this.mutationDispatcher.retain(style)
            : undefined;
          this.conditionalStyles.splice(i, 1);
          if (before) {
            this.mutationDispatcher.report([
              {
                kind: "conditional-style",
                before: { index: i, value: before },
              },
            ]);
          }
          return true;
        }
        currentIndex++;
      }
    }
    return false;
  }

  /**
   * Get all conditional styles intersecting with a range
   */
  getConditionalStylesIntersectingWithRange(
    range: RangeAddress
  ): ConditionalStyle[] {
    return this.conditionalStyles.filter((style) =>
      style.areas.some(
        (area) =>
          area.workbookName === range.workbookName &&
          area.sheetName === range.sheetName &&
          rangesIntersect(area.range, range.range)
      )
    );
  }

  /**
   * Add a direct cell style rule
   */
  addCellStyle(style: DirectCellStyle): void {
    const normalizedStyle = normalizeDirectCellStyle(style);
    const index = this.cellStyles.length;
    const after = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(normalizedStyle)
      : undefined;
    this.cellStyles.push(normalizedStyle);
    if (after) {
      this.mutationDispatcher.report([
        { kind: "cell-style", after: { index, value: after } },
      ]);
    }
  }

  addCellDataType(dataType: DirectCellDataType): void {
    const index = this.cellDataTypes.length;
    const after = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(dataType)
      : undefined;
    this.cellDataTypes.push(dataType);
    if (after) {
      this.mutationDispatcher.report([
        { kind: "cell-data-type", after: { index, value: after } },
      ]);
    }
  }

  getAllCellDataTypes(): DirectCellDataType[] {
    return [...this.cellDataTypes];
  }

  getCellDataTypesIntersectingWithRange(
    range: RangeAddress
  ): DirectCellDataType[] {
    return this.cellDataTypes.filter((rule) =>
      rule.areas.some(
        (area) =>
          area.workbookName === range.workbookName &&
          area.sheetName === range.sheetName &&
          rangesIntersect(area.range, range.range)
      )
    );
  }

  getCellDataType(cellAddress: CellAddress): CellDataType {
    for (let index = this.cellDataTypes.length - 1; index >= 0; index--) {
      const rule = this.cellDataTypes[index];
      if (
        rule?.areas.some(
          (area) =>
            area.workbookName === cellAddress.workbookName &&
            area.sheetName === cellAddress.sheetName &&
            isCellInRange(cellAddress, area.range)
        )
      ) {
        return rule.dataType;
      }
    }

    return "general";
  }

  getDataTypeForRange(range: RangeAddress): CellDataType | undefined {
    let unresolved = [range.range];
    let resolvedType: CellDataType | undefined;

    for (let index = this.cellDataTypes.length - 1; index >= 0; index--) {
      const rule = this.cellDataTypes[index];
      if (!rule) continue;

      for (const area of rule.areas) {
        if (
          area.workbookName !== range.workbookName ||
          area.sheetName !== range.sheetName
        ) {
          continue;
        }

        const nextUnresolved = [];
        let covered = false;
        for (const remaining of unresolved) {
          const intersection = intersectRanges(remaining, area.range);
          if (!intersection) {
            nextUnresolved.push(remaining);
            continue;
          }

          covered = true;
          nextUnresolved.push(...subtractRange(remaining, intersection));
        }

        if (covered) {
          if (resolvedType !== undefined && resolvedType !== rule.dataType) {
            return undefined;
          }
          resolvedType = rule.dataType;
          unresolved = nextUnresolved;
          if (unresolved.length === 0) {
            return resolvedType;
          }
        }
      }
    }

    if (unresolved.length > 0) {
      if (resolvedType !== undefined && resolvedType !== "general") {
        return undefined;
      }
      return "general";
    }

    return resolvedType ?? "general";
  }

  /**
   * Remove a direct cell style rule by index for a specific workbook
   */
  removeCellStyle(workbookName: string, index: number): boolean {
    const workbookStyles = this.cellStyles.filter(
      (style) =>
        style &&
        style.areas &&
        style.areas.some((area) => area.workbookName === workbookName)
    );
    if (index < 0 || index >= workbookStyles.length) {
      return false;
    }
    // Find the actual index in the full array
    let currentIndex = 0;
    for (let i = 0; i < this.cellStyles.length; i++) {
      const style = this.cellStyles[i];
      if (
        style &&
        style.areas &&
        style.areas.some((area) => area.workbookName === workbookName)
      ) {
        if (currentIndex === index) {
          const before = this.mutationDispatcher.observed
            ? this.mutationDispatcher.retain(style)
            : undefined;
          this.cellStyles.splice(i, 1);
          if (before) {
            this.mutationDispatcher.report([
              { kind: "cell-style", before: { index: i, value: before } },
            ]);
          }
          return true;
        }
        currentIndex++;
      }
    }
    return false;
  }

  /**
   * Get all direct cell styles intersecting with a range
   */
  getStylesIntersectingWithRange(range: RangeAddress): DirectCellStyle[] {
    return this.cellStyles.filter(
      (style) =>
        style &&
        style.areas.some(
          (area) =>
            area.sheetName === range.sheetName &&
            area.workbookName === range.workbookName &&
            rangesIntersect(area.range, range.range)
        )
    );
  }

  /**
   * Get the style for a range if all cells in the range have the same style
   * Returns the DirectCellStyle if the range is completely contained within a single style's areas
   * Returns undefined if multiple styles, partial coverage, or no styles apply
   */
  getStyleForRange(range: RangeAddress): DirectCellStyle | undefined {
    const intersectingStyles = this.getStylesIntersectingWithRange(range);

    // If no styles intersect, return undefined
    if (intersectingStyles.length === 0) {
      return undefined;
    }

    // If multiple styles intersect, return undefined (range has mixed styles)
    if (intersectingStyles.length > 1) {
      return undefined;
    }

    // Check if the range is completely contained within any of the single style's areas
    const style = intersectingStyles[0]!;
    const isContained = style.areas.some(
      (area) =>
        area.workbookName === range.workbookName &&
        area.sheetName === range.sheetName &&
        isRangeContained(range.range, area.range)
    );

    if (isContained) {
      return style;
    }

    // Range is not completely contained, return undefined
    return undefined;
  }

  /**
   * Get all conditional styles across all workbooks (for serialization)
   */
  getAllConditionalStyles(): ConditionalStyle[] {
    return [...this.conditionalStyles];
  }

  /**
   * Get all cell styles (for serialization)
   */
  getAllCellStyles(): DirectCellStyle[] {
    return [...this.cellStyles];
  }

  /**
   * Reset all styles (for deserialization)
   */
  resetStyles(
    conditionalStyles?: ConditionalStyle[],
    cellStyles?: DirectCellStyle[],
    cellDataTypes?: DirectCellDataType[]
  ): void {
    this.observeCollections(
      ["conditional-style", "cell-style", "cell-data-type"],
      () => {
        this.conditionalStyles = conditionalStyles
          ? [...conditionalStyles]
          : [];
        this.cellStyles = cellStyles
          ? cellStyles.map(normalizeDirectCellStyle)
          : [];
        this.cellDataTypes = cellDataTypes ? [...cellDataTypes] : [];
      }
    );
  }

  toSnapshot(): StyleManagerSnapshot {
    return {
      conditionalStyles: this.getAllConditionalStyles(),
      cellStyles: this.getAllCellStyles(),
      cellDataTypes: this.getAllCellDataTypes(),
    };
  }

  restoreFromSnapshot(snapshot: StyleManagerSnapshot): void {
    this.resetStyles(
      snapshot.conditionalStyles,
      snapshot.cellStyles,
      snapshot.cellDataTypes
    );
  }

  /**
   * Remove all styles for a workbook
   */
  removeWorkbookStyles(workbookName: string): void {
    this.observeCollections(
      ["conditional-style", "cell-style", "cell-data-type"],
      () => {
        this.conditionalStyles = this.conditionalStyles.filter(
          (style) =>
            !style.areas.some((area) => area.workbookName === workbookName)
        );
        this.cellStyles = this.cellStyles.filter(
          (style) =>
            !style.areas.some((area) => area.workbookName === workbookName)
        );
        this.cellDataTypes = this.cellDataTypes
          .map((rule) => {
            if (
              !rule.areas.some((area) => area.workbookName === workbookName)
            ) {
              return rule;
            }
            return {
              ...rule,
              areas: rule.areas.filter(
                (area) => area.workbookName !== workbookName
              ),
            };
          })
          .filter((rule) => rule.areas.length > 0);
      }
    );
  }

  /**
   * Update workbook name in all style references
   */
  updateWorkbookName(oldName: string, newName: string): void {
    if (oldName === newName) {
      return;
    }
    this.observeCollections(
      ["conditional-style", "cell-style", "cell-data-type"],
      () => {
        this.conditionalStyles = this.conditionalStyles.map((style) =>
          this.renameWorkbookInRule(style, oldName, newName)
        );
        this.cellStyles = this.cellStyles.map((style) =>
          this.renameWorkbookInRule(style, oldName, newName)
        );
        this.cellDataTypes = this.cellDataTypes.map((rule) =>
          this.renameWorkbookInRule(rule, oldName, newName)
        );
      }
    );
  }

  /**
   * Update sheet name in style references
   */
  updateSheetName(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    if (oldSheetName === newSheetName) {
      return;
    }
    this.observeCollections(
      ["conditional-style", "cell-style", "cell-data-type"],
      () => {
        this.conditionalStyles = this.conditionalStyles.map((style) =>
          this.renameSheetInRule(
            style,
            workbookName,
            oldSheetName,
            newSheetName
          )
        );
        this.cellStyles = this.cellStyles.map((style) =>
          this.renameSheetInRule(
            style,
            workbookName,
            oldSheetName,
            newSheetName
          )
        );
        this.cellDataTypes = this.cellDataTypes.map((rule) =>
          this.renameSheetInRule(rule, workbookName, oldSheetName, newSheetName)
        );
      }
    );
  }

  /**
   * Remove styles that reference a deleted sheet
   */
  removeSheetStyles(workbookName: string, sheetName: string): void {
    this.observeCollections(
      ["conditional-style", "cell-style", "cell-data-type"],
      () => {
        this.conditionalStyles = this.conditionalStyles.filter(
          (style) =>
            !style.areas.some(
              (area) =>
                area.workbookName === workbookName &&
                area.sheetName === sheetName
            )
        );
        this.cellStyles = this.cellStyles.filter(
          (style) =>
            !style.areas.some(
              (area) =>
                area.workbookName === workbookName &&
                area.sheetName === sheetName
            )
        );
        this.cellDataTypes = this.cellDataTypes
          .map((rule) => {
            if (
              !rule.areas.some(
                (area) =>
                  area.workbookName === workbookName &&
                  area.sheetName === sheetName
              )
            ) {
              return rule;
            }
            return {
              ...rule,
              areas: rule.areas.filter(
                (area) =>
                  area.workbookName !== workbookName ||
                  area.sheetName !== sheetName
              ),
            };
          })
          .filter((rule) => rule.areas.length > 0);
      }
    );
  }

  /**
   * Get the style for a specific cell.
   * Direct cell styles compose in insertion order, with later styles overriding
   * earlier styles for the same properties. Conditional styles then layer over
   * direct styles for the properties they define.
   */
  getCellStyle(cellAddress: CellAddress): CellStyle | undefined {
    let resolvedStyle: CellStyle | undefined;

    for (const cellStyle of this.cellStyles) {
      if (!cellStyle || !cellStyle.areas) {
        continue;
      }

      for (const area of cellStyle.areas) {
        if (
          area.workbookName === cellAddress.workbookName &&
          area.sheetName === cellAddress.sheetName &&
          isCellInRange(cellAddress, area.range)
        ) {
          resolvedStyle = {
            ...resolvedStyle,
            ...cellStyle.style,
          };
          break;
        }
      }
    }

    for (const style of this.conditionalStyles) {
      if (!style || !style.areas) {
        continue;
      }

      // Check if cell is in any of the style's areas
      for (const area of style.areas) {
        if (
          area.sheetName !== cellAddress.sheetName ||
          area.workbookName !== cellAddress.workbookName
        ) {
          continue;
        }

        if (!isCellInRange(cellAddress, area.range)) {
          continue;
        }

        // Cell is in area, evaluate condition
        if (style.condition.type === "formula") {
          const result = this.evaluateFormulaCondition(
            cellAddress,
            style,
            area
          );
          if (result) return { ...resolvedStyle, ...result };
        } else {
          const result = this.evaluateGradientCondition(
            cellAddress,
            style,
            area
          );
          if (result) return { ...resolvedStyle, ...result };
        }
      }
    }

    return resolvedStyle;
  }

  /**
   * Evaluate a formula-based style condition
   */
  private evaluateFormulaCondition(
    cellAddress: CellAddress,
    style: ConditionalStyle,
    area: RangeAddress
  ): CellStyle | undefined {
    if (style.condition.type !== "formula") {
      return undefined;
    }

    try {
      // Evaluate formula in context of the cell
      // evaluateFormula expects a full cell value (with = prefix for formulas)
      const formula = style.condition.formula.startsWith("=")
        ? style.condition.formula
        : `=${style.condition.formula}`;

      const result = this.evaluationManager.evaluateFormula(
        formula,
        cellAddress
      );

      // Check if result is truthy
      const isTruthy =
        result === true ||
        result === "TRUE" ||
        (typeof result === "number" && result !== 0);

      if (isTruthy) {
        return {
          backgroundColor: lchToHex(style.condition.color),
        };
      }
    } catch (error) {
      // If formula evaluation fails, don't apply style
      console.warn("Failed to evaluate formula condition:", error);
    }

    return undefined;
  }

  /**
   * Evaluate a gradient-based style condition
   */
  private evaluateGradientCondition(
    cellAddress: CellAddress,
    style: ConditionalStyle,
    area: RangeAddress
  ): CellStyle | undefined {
    if (style.condition.type !== "gradient") {
      return undefined;
    }

    try {
      // Get the cell's evaluation result
      const evalResult =
        this.evaluationManager.getCellEvaluationResult(cellAddress);
      if (!evalResult || evalResult.type !== "value") {
        return undefined;
      }
      if (evalResult.result.type !== "number") {
        return undefined;
      }
      const cellValue = evalResult.result.value;

      // Calculate min and max values for the gradient
      const { min: minValue, max: maxValue } = this.calculateGradientBounds(
        style,
        cellAddress,
        area
      );

      if (minValue === null || maxValue === null) {
        return undefined;
      }

      // Calculate interpolation factor
      const factor = calculateGradientFactor(cellValue, minValue, maxValue);

      // Interpolate between min and max colors
      const minColor = style.condition.min.color;
      const maxColor = style.condition.max.color;
      const interpolatedColor = interpolateLCH(minColor, maxColor, factor);

      return {
        backgroundColor: lchToHex(interpolatedColor),
      };
    } catch (error) {
      console.warn("Failed to evaluate gradient condition:", error);
      return undefined;
    }
  }

  /**
   * Calculate min and max bounds for a gradient
   */
  private calculateGradientBounds(
    style: ConditionalStyle,
    cellAddress: CellAddress,
    area: RangeAddress
  ): { min: number | null; max: number | null } {
    if (style.condition.type !== "gradient") {
      return { min: null, max: null };
    }

    const { min: minConfig, max: maxConfig } = style.condition;
    const topLeftCell: CellAddress = {
      workbookName: area.workbookName,
      sheetName: area.sheetName,
      colIndex: area.range.start.col,
      rowIndex: area.range.start.row,
    };

    // Calculate min value
    let minValue: number | null = null;
    if (minConfig.type === "lowest_value") {
      // Evaluate MIN(range) formula directly
      try {
        const rangeRef = this.getRangeReference(area);
        const result = this.evaluationManager.evaluateFormula(
          `=MIN(${rangeRef})`,
          topLeftCell
        );
        if (typeof result === "number") {
          minValue = result;
        }
      } catch (error) {
        console.warn("Failed to calculate MIN:", error);
      }
    } else {
      // Evaluate valueFormula in context of area's top-left cell
      const formula = minConfig.valueFormula.startsWith("=")
        ? minConfig.valueFormula
        : `=${minConfig.valueFormula}`;
      const result = this.evaluationManager.evaluateFormula(
        formula,
        topLeftCell
      );
      if (typeof result === "number") {
        minValue = result;
      }
    }

    // Calculate max value
    let maxValue: number | null = null;
    if (maxConfig.type === "highest_value") {
      // Evaluate MAX(range) formula directly
      try {
        const rangeRef = this.getRangeReference(area);
        const result = this.evaluationManager.evaluateFormula(
          `=MAX(${rangeRef})`,
          topLeftCell
        );
        if (typeof result === "number") {
          maxValue = result;
        }
      } catch (error) {
        console.warn("Failed to calculate MAX:", error);
      }
    } else {
      // Evaluate valueFormula in context of area's top-left cell
      const formula = maxConfig.valueFormula.startsWith("=")
        ? maxConfig.valueFormula
        : `=${maxConfig.valueFormula}`;
      const result = this.evaluationManager.evaluateFormula(
        formula,
        topLeftCell
      );
      if (typeof result === "number") {
        maxValue = result;
      }
    }

    return { min: minValue, max: maxValue };
  }

  /**
   * Get a range reference string from a RangeAddress
   * Follows CANONICAL_RANGES.md format:
   * - Closed: A5:D10
   * - Row-bounded (col-open): A5:10
   * - Col-bounded (row-open): A5:D
   * - Open both: A5:INFINITY
   */
  private getRangeReference(area: RangeAddress): string {
    const colToLetter = (col: number): string => {
      let result = "";
      let c = col;
      while (c >= 0) {
        result = String.fromCharCode(65 + (c % 26)) + result;
        c = Math.floor(c / 26) - 1;
      }
      return result;
    };

    const startCol = colToLetter(area.range.start.col);
    const startRow = area.range.start.row + 1; // Convert to 1-based

    const isColInfinity = area.range.end.col.type === "infinity";
    const isRowInfinity = area.range.end.row.type === "infinity";

    let rangeStr: string;

    if (isColInfinity && isRowInfinity) {
      // Open both: A5:INFINITY
      rangeStr = `${startCol}${startRow}:INFINITY`;
    } else if (isColInfinity) {
      // Row-bounded (col-open): A5:10
      if (area.range.end.row.type === "number") {
        const endRow = area.range.end.row.value + 1; // Convert to 1-based
        rangeStr = `${startCol}${startRow}:${endRow}`;
      } else {
        rangeStr = `${startCol}${startRow}:INFINITY`;
      }
    } else if (isRowInfinity) {
      // Col-bounded (row-open): A5:D
      if (area.range.end.col.type === "number") {
        const endCol = colToLetter(area.range.end.col.value);
        rangeStr = `${startCol}${startRow}:${endCol}`;
      } else {
        rangeStr = `${startCol}${startRow}:INFINITY`;
      }
    } else {
      // Closed rectangle: A5:D10
      if (
        area.range.end.col.type === "number" &&
        area.range.end.row.type === "number"
      ) {
        const endCol = colToLetter(area.range.end.col.value);
        const endRow = area.range.end.row.value + 1; // Convert to 1-based
        rangeStr = `${startCol}${startRow}:${endCol}${endRow}`;
      } else {
        // Fallback to INFINITY if types don't match
        rangeStr = `${startCol}${startRow}:INFINITY`;
      }
    }

    // Quote sheet name if it contains spaces or special characters
    const needsQuotes = /[ '!]/.test(area.sheetName);
    const sheetRef = needsQuotes
      ? `'${area.sheetName.replace(/'/g, "''")}'`
      : area.sheetName;

    // Construct the full reference: [workbook]'sheet'!range
    return `[${area.workbookName}]${sheetRef}!${rangeStr}`;
  }

  /**
   * Clear cell styles and conditional styles for a given range
   * Adjusts existing style ranges rather than deleting them entirely
   */
  clearCellStyles(range: RangeAddress): void {
    this.observeCollections(["cell-style", "conditional-style"], () => {
      this.cellStyles = this.cellStyles
        .map((style) => this.subtractRuleRange(style, range))
        .filter((style): style is DirectCellStyle => style !== undefined);
      this.conditionalStyles = this.conditionalStyles
        .map((style) => this.subtractRuleRange(style, range))
        .filter((style): style is ConditionalStyle => style !== undefined);
    });
  }

  /**
   * Clear cell styles in a range using subtraction
   * For each intersecting style, subtract the cleared range from its areas:
   * - If an area is completely contained: remove that area
   * - If an area partially overlaps: split into remaining rectangles (hole punching)
   * - If no intersection: keep area unchanged
   *
   * This matches Excel's behavior where cutting/pasting creates multi-area styles
   */
  clearCellStylesInRange(range: RangeAddress): void {
    this.observeCollections(["cell-style"], () => {
      this.cellStyles = this.cellStyles
        .map((style) => this.subtractRuleRange(style, range))
        .filter((style): style is DirectCellStyle => style !== undefined);
    });
  }

  clearCellDataTypesInRange(range: RangeAddress): void {
    this.observeCollections(["cell-data-type"], () => {
      this.cellDataTypes = this.cellDataTypes
        .map((rule) => this.subtractRuleRange(rule, range))
        .filter((rule): rule is DirectCellDataType => rule !== undefined);
    });
  }

  /** Applies retained deltas directly without notifying the observer. */
  applyHistoryChanges(
    changes: readonly StyleDataChange[],
    direction: MutationDirection
  ): void {
    const conditionalChanges = changes.filter(
      (change): change is ConditionalStyleDataChange =>
        change.kind === "conditional-style"
    );
    const cellStyleChanges = changes.filter(
      (change): change is CellStyleDataChange => change.kind === "cell-style"
    );
    const dataTypeChanges = changes.filter(
      (change): change is CellDataTypeDataChange =>
        change.kind === "cell-data-type"
    );

    this.conditionalStyles = applyIndexedChanges(
      this.conditionalStyles,
      conditionalChanges,
      direction
    );
    this.cellStyles = applyIndexedChanges(
      this.cellStyles,
      cellStyleChanges,
      direction
    );
    this.cellDataTypes = applyIndexedChanges(
      this.cellDataTypes,
      dataTypeChanges,
      direction
    );
  }

  private observeCollections(
    kinds: readonly StyleCollectionKind[],
    callback: () => void
  ): void {
    if (!this.mutationDispatcher.observed) {
      callback();
      return;
    }

    const beforeConditional = kinds.includes("conditional-style")
      ? [...this.conditionalStyles]
      : undefined;
    const beforeCellStyles = kinds.includes("cell-style")
      ? [...this.cellStyles]
      : undefined;
    const beforeDataTypes = kinds.includes("cell-data-type")
      ? [...this.cellDataTypes]
      : undefined;

    callback();

    const changes: StyleDataChange[] = [];
    if (beforeConditional) {
      changes.push(
        ...this.diffCollection(
          "conditional-style",
          beforeConditional,
          this.conditionalStyles
        )
      );
    }
    if (beforeCellStyles) {
      changes.push(
        ...this.diffCollection("cell-style", beforeCellStyles, this.cellStyles)
      );
    }
    if (beforeDataTypes) {
      changes.push(
        ...this.diffCollection(
          "cell-data-type",
          beforeDataTypes,
          this.cellDataTypes
        )
      );
    }
    this.mutationDispatcher.report(changes);
  }

  private diffCollection<TValue>(
    kind: StyleCollectionKind,
    before: readonly TValue[],
    after: readonly TValue[]
  ): StyleDataChange[] {
    const afterPositions = new Map<TValue, number[]>();
    for (let index = 0; index < after.length; index++) {
      const value = after[index]!;
      const positions = afterPositions.get(value);
      if (positions) {
        positions.push(index);
      } else {
        afterPositions.set(value, [index]);
      }
    }

    const nextPosition = new Map<TValue, number>();
    const matches: Array<{ beforeIndex: number; afterIndex: number }> = [];
    for (let beforeIndex = 0; beforeIndex < before.length; beforeIndex++) {
      const value = before[beforeIndex]!;
      const occurrence = nextPosition.get(value) ?? 0;
      const afterIndex = afterPositions.get(value)?.[occurrence];
      if (afterIndex !== undefined) {
        matches.push({ beforeIndex, afterIndex });
        nextPosition.set(value, occurrence + 1);
      }
    }

    // An identity can remain in place iff it belongs to an increasing
    // subsequence of target indexes. Everything else is represented as one
    // sparse removal/insertion pair, preserving precedence exactly.
    const tails: number[] = [];
    const tailMatchIndexes: number[] = [];
    const predecessors = new Array<number>(matches.length).fill(-1);
    for (let matchIndex = 0; matchIndex < matches.length; matchIndex++) {
      const afterIndex = matches[matchIndex]!.afterIndex;
      let low = 0;
      let high = tails.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        if (tails[middle]! < afterIndex) {
          low = middle + 1;
        } else {
          high = middle;
        }
      }
      if (low > 0) {
        predecessors[matchIndex] = tailMatchIndexes[low - 1]!;
      }
      tails[low] = afterIndex;
      tailMatchIndexes[low] = matchIndex;
    }

    const keptBeforeIndexes = new Set<number>();
    const keptAfterIndexes = new Set<number>();
    let keptMatchIndex = tailMatchIndexes[tails.length - 1] ?? -1;
    while (keptMatchIndex >= 0) {
      const match = matches[keptMatchIndex]!;
      keptBeforeIndexes.add(match.beforeIndex);
      keptAfterIndexes.add(match.afterIndex);
      keptMatchIndex = predecessors[keptMatchIndex]!;
    }

    const changes: StyleDataChange[] = [];
    for (let index = 0; index < before.length; index++) {
      if (!keptBeforeIndexes.has(index)) {
        changes.push({
          kind,
          before: {
            index,
            value: this.mutationDispatcher.retain(before[index]!),
          },
        } as unknown as StyleDataChange);
      }
    }
    for (let index = 0; index < after.length; index++) {
      if (!keptAfterIndexes.has(index)) {
        changes.push({
          kind,
          after: {
            index,
            value: this.mutationDispatcher.retain(after[index]!),
          },
        } as unknown as StyleDataChange);
      }
    }
    return changes;
  }

  private renameWorkbookInRule<TValue extends { areas: RangeAddress[] }>(
    value: TValue,
    oldName: string,
    newName: string
  ): TValue {
    if (!value.areas.some((area) => area.workbookName === oldName)) {
      return value;
    }
    return {
      ...value,
      areas: value.areas.map((area) =>
        area.workbookName === oldName
          ? { ...area, workbookName: newName }
          : area
      ),
    };
  }

  private renameSheetInRule<TValue extends { areas: RangeAddress[] }>(
    value: TValue,
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): TValue {
    if (
      !value.areas.some(
        (area) =>
          area.workbookName === workbookName && area.sheetName === oldSheetName
      )
    ) {
      return value;
    }
    return {
      ...value,
      areas: value.areas.map((area) =>
        area.workbookName === workbookName && area.sheetName === oldSheetName
          ? { ...area, sheetName: newSheetName }
          : area
      ),
    };
  }

  private subtractRuleRange<TValue extends { areas: RangeAddress[] }>(
    value: TValue,
    range: RangeAddress
  ): TValue | undefined {
    if (
      !value.areas.some(
        (area) =>
          area.workbookName === range.workbookName &&
          area.sheetName === range.sheetName &&
          rangesIntersect(area.range, range.range)
      )
    ) {
      return value;
    }

    const areas = value.areas.flatMap((area) => {
      if (
        area.workbookName !== range.workbookName ||
        area.sheetName !== range.sheetName ||
        !rangesIntersect(area.range, range.range)
      ) {
        return [area];
      }
      return subtractRange(area.range, range.range).map(
        (remainingRange): RangeAddress => ({
          workbookName: area.workbookName,
          sheetName: area.sheetName,
          range: remainingRange,
        })
      );
    });

    return areas.length > 0 ? { ...value, areas } : undefined;
  }
}
