/**
 * StyleManager - Manages conditional styling for cells
 */

import type {
  CellAddress,
  CellStyle,
  ConditionalStyle,
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
} from "../utils/range-utils";

export class StyleManager {
  private conditionalStyles: ConditionalStyle[] = [];
  private cellStyles: DirectCellStyle[] = [];

  constructor(private evaluationManager: EvaluationManager) {}

  /**
   * Add a conditional style rule
   */
  addConditionalStyle(style: ConditionalStyle): void {
    this.conditionalStyles.push(style);
  }

  /**
   * Remove a conditional style rule by index for a specific workbook
   */
  removeConditionalStyle(workbookName: string, index: number): boolean {
    const workbookStyles = this.conditionalStyles.filter(
      (style) => style.areas.some(area => area.workbookName === workbookName)
    );
    if (index < 0 || index >= workbookStyles.length) {
      return false;
    }
    // Find the actual index in the full array
    let currentIndex = 0;
    for (let i = 0; i < this.conditionalStyles.length; i++) {
      const style = this.conditionalStyles[i];
      if (style && style.areas.some(area => area.workbookName === workbookName)) {
        if (currentIndex === index) {
          this.conditionalStyles.splice(i, 1);
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
    return this.conditionalStyles.filter(
      (style) =>
        style.areas.some(area =>
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
    this.cellStyles.push(style);
  }

  /**
   * Remove a direct cell style rule by index for a specific workbook
   */
  removeCellStyle(workbookName: string, index: number): boolean {
    const workbookStyles = this.cellStyles.filter(
      (style) => style && style.areas && style.areas.some(area => area.workbookName === workbookName)
    );
    if (index < 0 || index >= workbookStyles.length) {
      return false;
    }
    // Find the actual index in the full array
    let currentIndex = 0;
    for (let i = 0; i < this.cellStyles.length; i++) {
      const style = this.cellStyles[i];
      if (style && style.areas && style.areas.some(area => area.workbookName === workbookName)) {
        if (currentIndex === index) {
          this.cellStyles.splice(i, 1);
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
        style.areas.some(area =>
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
    const isContained = style.areas.some(area =>
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
    cellStyles?: DirectCellStyle[]
  ): void {
    this.conditionalStyles = conditionalStyles ? [...conditionalStyles] : [];
    this.cellStyles = cellStyles ? [...cellStyles] : [];
  }

  toSnapshot(): StyleManagerSnapshot {
    return {
      conditionalStyles: this.getAllConditionalStyles(),
      cellStyles: this.getAllCellStyles(),
    };
  }

  restoreFromSnapshot(snapshot: StyleManagerSnapshot): void {
    this.resetStyles(snapshot.conditionalStyles, snapshot.cellStyles);
  }

  /**
   * Remove all styles for a workbook
   */
  removeWorkbookStyles(workbookName: string): void {
    this.conditionalStyles = this.conditionalStyles.filter(
      (style) => !style.areas.some(area => area.workbookName === workbookName)
    );
    this.cellStyles = this.cellStyles.filter(
      (style) => !style.areas.some(area => area.workbookName === workbookName)
    );
  }

  /**
   * Update workbook name in all style references
   */
  updateWorkbookName(oldName: string, newName: string): void {
    // Update conditional styles
    this.conditionalStyles = this.conditionalStyles.map((style) => ({
      ...style,
      areas: style.areas.map(area =>
        area.workbookName === oldName
          ? { ...area, workbookName: newName }
          : area
      )
    }));
    // Update cell styles
    this.cellStyles = this.cellStyles.map((style) => ({
      ...style,
      areas: style.areas.map(area =>
        area.workbookName === oldName
          ? { ...area, workbookName: newName }
          : area
      )
    }));
  }

  /**
   * Update sheet name in style references
   */
  updateSheetName(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    // Update conditional styles
    this.conditionalStyles = this.conditionalStyles.map((style) => ({
      ...style,
      areas: style.areas.map(area =>
        area.workbookName === workbookName && area.sheetName === oldSheetName
          ? { ...area, sheetName: newSheetName }
          : area
      )
    }));
    // Update cell styles
    this.cellStyles = this.cellStyles.map((style) => ({
      ...style,
      areas: style.areas.map(area =>
        area.workbookName === workbookName && area.sheetName === oldSheetName
          ? { ...area, sheetName: newSheetName }
          : area
      )
    }));
  }

  /**
   * Remove styles that reference a deleted sheet
   */
  removeSheetStyles(workbookName: string, sheetName: string): void {
    this.conditionalStyles = this.conditionalStyles.filter(
      (style) =>
        !style.areas.some(area =>
          area.workbookName === workbookName &&
          area.sheetName === sheetName
        )
    );
    this.cellStyles = this.cellStyles.filter(
      (style) =>
        !style.areas.some(area =>
          area.workbookName === workbookName &&
          area.sheetName === sheetName
        )
    );
  }

  /**
   * Get the style for a specific cell
   * Returns the first matching style (first match wins)
   * Checks cellStyles first, then conditionalStyles
   */
  getCellStyle(cellAddress: CellAddress): CellStyle | undefined {
    // First check conditional styles
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
          const result = this.evaluateFormulaCondition(cellAddress, style, area);
          if (result) return result;
        } else {
          const result = this.evaluateGradientCondition(cellAddress, style, area);
          if (result) return result;
        }
      }
    }

    // Then check direct cell styles
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
          return cellStyle.style;
        }
      }
    }

    return undefined;
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
      const { min: minValue, max: maxValue} = this.calculateGradientBounds(
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
    // Process cellStyles - punch holes in areas
    this.cellStyles = this.cellStyles.map(cellStyle => {
      if (!cellStyle || !cellStyle.areas) {
        return cellStyle;
      }

      const newAreas: RangeAddress[] = [];
      
      for (const area of cellStyle.areas) {
        // Check if this area intersects with the clear range
        if (
          area.workbookName === range.workbookName &&
          area.sheetName === range.sheetName &&
          rangesIntersect(area.range, range.range)
        ) {
          // Subtract the clear range from this area
          const remainingRanges = subtractRange(area.range, range.range);

          // Add all remaining ranges as new areas
          for (const remainingRange of remainingRanges) {
            newAreas.push({
              workbookName: area.workbookName,
              sheetName: area.sheetName,
              range: remainingRange,
            });
          }
        } else {
          // No intersection, keep the area as-is
          newAreas.push(area);
        }
      }
      
      return { ...cellStyle, areas: newAreas };
    }).filter(style => style.areas.length > 0);

    // Process conditionalStyles - punch holes in areas
    this.conditionalStyles = this.conditionalStyles.map(conditionalStyle => {
      if (!conditionalStyle || !conditionalStyle.areas) {
        return conditionalStyle;
      }

      const newAreas: RangeAddress[] = [];
      
      for (const area of conditionalStyle.areas) {
        // Check if this area intersects with the clear range
        if (
          area.workbookName === range.workbookName &&
          area.sheetName === range.sheetName &&
          rangesIntersect(area.range, range.range)
        ) {
          // Subtract the clear range from this area
          const remainingRanges = subtractRange(area.range, range.range);

          // Add all remaining ranges as new areas
          for (const remainingRange of remainingRanges) {
            newAreas.push({
              workbookName: area.workbookName,
              sheetName: area.sheetName,
              range: remainingRange,
            });
          }
        } else {
          // No intersection, keep the area as-is
          newAreas.push(area);
        }
      }
      
      return { ...conditionalStyle, areas: newAreas };
    }).filter(style => style.areas.length > 0);
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
    this.cellStyles = this.cellStyles.map(style => {
      const newAreas: RangeAddress[] = [];
      
      for (const area of style.areas) {
        // Skip areas from different sheets/workbooks
        if (
          area.workbookName !== range.workbookName ||
          area.sheetName !== range.sheetName
        ) {
          newAreas.push(area);
          continue;
        }

        // Check if this area intersects with the range to clear
        if (!rangesIntersect(area.range, range.range)) {
          // No intersection, keep the area unchanged
          newAreas.push(area);
          continue;
        }

        // Area intersects - subtract the cleared range (may produce multiple ranges)
        const remainingRanges = subtractRange(area.range, range.range);

        // Add all remaining ranges as new areas for this style
        for (const remainingRange of remainingRanges) {
          newAreas.push({
            workbookName: area.workbookName,
            sheetName: area.sheetName,
            range: remainingRange,
          });
        }
      }
      
      return { ...style, areas: newAreas };
    }).filter(style => style.areas.length > 0); // Remove styles with no areas left
  }
}
