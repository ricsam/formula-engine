import { describe, it, expect, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { RangeAddress } from "../../../src/core/types";

interface TestRangeMetadata {
  kind: "validation" | "annotation";
  label: string;
}

const workbookName = "wb1";
const sheetName = "sheet1";

const range = (
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number
): RangeAddress => ({
  workbookName,
  sheetName,
  range: {
    start: { col: startCol, row: startRow },
    end: {
      col: { type: "number", value: endCol },
      row: { type: "number", value: endRow },
    },
  },
});

describe("Range Metadata", () => {
  let engine: FormulaEngine<{ range: TestRangeMetadata }>;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty<{ range: TestRangeMetadata }>();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  it("sets, reads, and removes range metadata", () => {
    const id = engine.addRangeMetadata({
      areas: [range(0, 0, 1, 1)],
      metadata: { kind: "validation", label: "positive number" },
    });

    expect(engine.getAllRangeMetadata()).toHaveLength(1);
    expect(
      engine.getRangeMetadataForCell({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 1,
      })[0]?.metadata
    ).toEqual({ kind: "validation", label: "positive number" });
    expect(engine.getRangeMetadataIntersectingWithRange(range(1, 1, 2, 2))).toHaveLength(1);

    engine.removeRangeMetadata(id);
    expect(engine.getAllRangeMetadata()).toHaveLength(0);
  });

  it("clears only the overlapping portion of range metadata", () => {
    engine.addRangeMetadata({
      areas: [range(0, 0, 2, 0)],
      metadata: { kind: "annotation", label: "header" },
    });

    engine.clearRangeMetadata(range(1, 0, 1, 0));

    expect(engine.getRangeMetadataForCell({ workbookName, sheetName, colIndex: 0, rowIndex: 0 })).toHaveLength(1);
    expect(engine.getRangeMetadataForCell({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })).toHaveLength(0);
    expect(engine.getRangeMetadataForCell({ workbookName, sheetName, colIndex: 2, rowIndex: 0 })).toHaveLength(1);
  });

  it("copies range metadata when requested", () => {
    engine.addRangeMetadata({
      areas: [range(0, 0, 1, 0)],
      metadata: { kind: "annotation", label: "copy me" },
    });

    engine.pasteCells(
      [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      ],
      { workbookName, sheetName, colIndex: 3, rowIndex: 2 },
      { include: ["rangeMetadata"] }
    );

    const copied = engine.getRangeMetadataForCell({
      workbookName,
      sheetName,
      colIndex: 4,
      rowIndex: 2,
    });
    expect(copied).toHaveLength(1);
    expect(copied[0]?.metadata).toEqual({ kind: "annotation", label: "copy me" });
  });

  it("does not copy range metadata when only styles are requested", () => {
    engine.addRangeMetadata({
      areas: [range(0, 0, 0, 0)],
      metadata: { kind: "annotation", label: "do not copy" },
    });

    engine.pasteCells(
      [{ workbookName, sheetName, colIndex: 0, rowIndex: 0 }],
      { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      { include: ["style"] }
    );

    expect(engine.getRangeMetadataForCell({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })).toHaveLength(0);
  });

  it("serializes and deserializes range metadata", () => {
    engine.addRangeMetadata({
      areas: [range(0, 0, 1, 1)],
      metadata: { kind: "validation", label: "required" },
    });

    const restored = FormulaEngine.buildEmpty<{ range: TestRangeMetadata }>();
    restored.resetToSerializedEngine(engine.serializeEngine());

    expect(restored.getRangeMetadataForCell({ workbookName, sheetName, colIndex: 1, rowIndex: 1 })[0]?.metadata).toEqual({
      kind: "validation",
      label: "required",
    });
  });

  it("renames and clones range metadata with workbook and sheet changes", () => {
    engine.addRangeMetadata({
      areas: [range(0, 0, 0, 0)],
      metadata: { kind: "annotation", label: "moves" },
    });

    engine.renameSheet({ workbookName, sheetName, newSheetName: "renamed" });
    expect(
      engine.getRangeMetadataForCell({
        workbookName,
        sheetName: "renamed",
        colIndex: 0,
        rowIndex: 0,
      })[0]?.metadata
    ).toEqual({ kind: "annotation", label: "moves" });

    engine.cloneWorkbook(workbookName, "wb2");
    expect(
      engine.getRangeMetadataForCell({
        workbookName: "wb2",
        sheetName: "renamed",
        colIndex: 0,
        rowIndex: 0,
      })[0]?.metadata
    ).toEqual({ kind: "annotation", label: "moves" });
  });
});

describe("Cell Style extensions", () => {
  it("copies and serializes typed wrap and border style properties", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    engine.addCellStyle({
      areas: [range(0, 0, 0, 0)],
      style: { wrapText: true, borderColor: "#00FF00", borderSides: { left: true, top: true } },
    });

    engine.pasteCells(
      [{ workbookName, sheetName, colIndex: 0, rowIndex: 0 }],
      { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      { include: ["style"] }
    );

    const restored = FormulaEngine.buildEmpty();
    restored.resetToSerializedEngine(engine.serializeEngine());

    expect(restored.getCellStyle({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })?.wrapText).toBe(true);
    expect(restored.getCellStyle({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })?.borderColor).toBe("#00FF00");
    expect(restored.getCellStyle({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })?.borderSides).toEqual({
      left: true,
      top: true,
    });
  });
});
