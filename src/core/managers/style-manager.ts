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
import type { WorkbookManager } from "./workbook-manager";
import type { EvaluationManager } from "./evaluation-manager";
import { isCellInRange } from "../utils";
import {
  calculateGradientFactor,
  interpolateLCH,
  lchToHex,
} from "../utils/color-utils";
import { subtractRange, rangesIntersect, isRangeContained } from "../utils/range-utils";

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
      (style) => style.area.workbookName === workbookName
    );
    if (index < 0 || index >= workbookStyles.length) {
      return false;
    }
    // Find the actual index in the full array
    let currentIndex = 0;
    for (let i = 0; i < this.conditionalStyles.length; i++) {
      const style = this.conditionalStyles[i];
      if (style && style.area.workbookName === workbookName) {
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
  getConditionalStylesIntersectingWithRange(range: RangeAddress): ConditionalStyle[] {
    return this.conditionalStyles.filter(
      (style) =>
        style.area.workbookName === range.workbookName &&
        style.area.sheetName === range.sheetName &&
        rangesIntersect(style.area.range, range.range)
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
      (style) => style && style.area && style.area.workbookName === workbookName
    );
    if (index < 0 || index >= workbookStyles.length) {
      return false;
    }
    // Find the actual index in the full array
    let currentIndex = 0;
    for (let i = 0; i < this.cellStyles.length; i++) {
      const style = this.cellStyles[i];
      if (style && style.area && style.area.workbookName === workbookName) {
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
        style.area.sheetName === range.sheetName &&
        style.area.workbookName === range.workbookName &&
        rangesIntersect(style.area.range, range.range)
    );
  }

  /**
   * Get the style for a range if all cells in the range have the same style
   * Returns the DirectCellStyle if the range is completely contained within a single style's area
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
    
    // Check if the range is completely contained within the single style's area
    const style = intersectingStyles[0]!;
    if (isRangeContained(range.range, style.area.range)) {
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

  /**
   * Remove all styles for a workbook
   */
  removeWorkbookStyles(workbookName: string): void {
    this.conditionalStyles = this.conditionalStyles.filter(
      (style) => style.area.workbookName !== workbookName
    );
    this.cellStyles = this.cellStyles.filter(
      (style) => style.area.workbookName !== workbookName
    );
  }

  /**
   * Update workbook name in all style references
   */
  updateWorkbookName(oldName: string, newName: string): void {
    // Update conditional styles
    this.conditionalStyles = this.conditionalStyles.map((style) => {
      if (style.area.workbookName === oldName) {
        return {
          ...style,
          area: {
            ...style.area,
            workbookName: newName,
          },
        };
      }
      return style;
    });
    // Update cell styles
    this.cellStyles = this.cellStyles.map((style) => {
      if (style.area.workbookName === oldName) {
        return {
          ...style,
          area: {
            ...style.area,
            workbookName: newName,
          },
        };
      }
      return style;
    });
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
    this.conditionalStyles = this.conditionalStyles.map((style) => {
      if (
        style.area.workbookName === workbookName &&
        style.area.sheetName === oldSheetName
      ) {
        return {
          ...style,
          area: {
            ...style.area,
            sheetName: newSheetName,
          },
        };
      }
      return style;
    });
    // Update cell styles
    this.cellStyles = this.cellStyles.map((style) => {
      if (
        style.area.workbookName === workbookName &&
        style.area.sheetName === oldSheetName
      ) {
        return {
          ...style,
          area: {
            ...style.area,
            sheetName: newSheetName,
          },
        };
      }
      return style;
    });
  }

  /**
   * Remove styles that reference a deleted sheet
   */
  removeSheetStyles(workbookName: string, sheetName: string): void {
    this.conditionalStyles = this.conditionalStyles.filter(
      (style) =>
        !(
          style.area.workbookName === workbookName &&
          style.area.sheetName === sheetName
        )
    );
    this.cellStyles = this.cellStyles.filter(
      (style) =>
        !(
          style.area.workbookName === workbookName &&
          style.area.sheetName === sheetName
        )
    );
  }

  /**
   * Get the style for a specific cell
   * Returns the first matching style (first match wins)
   * Checks cellStyles first, then conditionalStyles
   */
  getCellStyle(cellAddress: CellAddress): CellStyle | undefined {
    // First check direct cell styles
    for (const cellStyle of this.cellStyles) {
      if (!cellStyle || !cellStyle.area) {
        continue;
      }
      if (
        cellStyle.area.workbookName === cellAddress.workbookName &&
        cellStyle.area.sheetName === cellAddress.sheetName &&
        isCellInRange(cellAddress, cellStyle.area.range)
      ) {
        return cellStyle.style;
      }
    }

    // Then check conditional styles
    for (const style of this.conditionalStyles) {
      if (!style || !style.area) {
        continue;
      }
      // Check if cell is in the style's area
      if (
        style.area.sheetName !== cellAddress.sheetName ||
        style.area.workbookName !== cellAddress.workbookName
      ) {
        continue;
      }

      if (!isCellInRange(cellAddress, style.area.range)) {
        continue;
      }

      // Cell is in area, evaluate condition
      if (style.condition.type === "formula") {
        const result = this.evaluateFormulaCondition(cellAddress, style);
        if (result) return result;
      } else {
        const result = this.evaluateGradientCondition(cellAddress, style);
        if (result) return result;
      }
    }

    return undefined;
  }

  /**
   * Evaluate a formula-based style condition
   */
  private evaluateFormulaCondition(
    cellAddress: CellAddress,
    style: ConditionalStyle
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
    style: ConditionalStyle
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
        cellAddress
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
    cellAddress: CellAddress
  ): { min: number | null; max: number | null } {
    if (style.condition.type !== "gradient") {
      return { min: null, max: null };
    }

    const { min: minConfig, max: maxConfig } = style.condition;
    const topLeftCell: CellAddress = {
      workbookName: style.area.workbookName,
      sheetName: style.area.sheetName,
      colIndex: style.area.range.start.col,
      rowIndex: style.area.range.start.row,
    };

    // Calculate min value
    let minValue: number | null = null;
    if (minConfig.type === "lowest_value") {
      // Evaluate MIN(range) formula directly
      try {
        const rangeRef = this.getRangeReference(style.area);
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
        const rangeRef = this.getRangeReference(style.area);
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
    // Process cellStyles
    const newCellStyles: DirectCellStyle[] = [];
    for (const cellStyle of this.cellStyles) {
      if (!cellStyle || !cellStyle.area) {
        newCellStyles.push(cellStyle);
        continue;
      }

      // Check if this style intersects with the clear range
      if (
        cellStyle.area.workbookName === range.workbookName &&
        cellStyle.area.sheetName === range.sheetName &&
        rangesIntersect(cellStyle.area.range, range.range)
      ) {
        // Subtract the clear range from this style's range
        const remainingRanges = subtractRange(
          cellStyle.area.range,
          range.range
        );

        // Add new styles for each remaining range
        for (const remainingRange of remainingRanges) {
          newCellStyles.push({
            area: {
              workbookName: cellStyle.area.workbookName,
              sheetName: cellStyle.area.sheetName,
              range: remainingRange,
            },
            style: cellStyle.style,
          });
        }
      } else {
        // No intersection, keep the style as-is
        newCellStyles.push(cellStyle);
      }
    }
    this.cellStyles = newCellStyles;

    // Process conditionalStyles
    const newConditionalStyles: ConditionalStyle[] = [];
    for (const conditionalStyle of this.conditionalStyles) {
      if (!conditionalStyle || !conditionalStyle.area) {
        newConditionalStyles.push(conditionalStyle);
        continue;
      }

      // Check if this style intersects with the clear range
      if (
        conditionalStyle.area.workbookName === range.workbookName &&
        conditionalStyle.area.sheetName === range.sheetName &&
        rangesIntersect(conditionalStyle.area.range, range.range)
      ) {
        // Subtract the clear range from this style's range
        const remainingRanges = subtractRange(
          conditionalStyle.area.range,
          range.range
        );

        // Add new styles for each remaining range
        for (const remainingRange of remainingRanges) {
          newConditionalStyles.push({
            area: {
              workbookName: conditionalStyle.area.workbookName,
              sheetName: conditionalStyle.area.sheetName,
              range: remainingRange,
            },
            condition: conditionalStyle.condition,
          });
        }
      } else {
        // No intersection, keep the style as-is
        newConditionalStyles.push(conditionalStyle);
      }
    }
    this.conditionalStyles = newConditionalStyles;
  }
}
