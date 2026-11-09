import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import {
  FormulaError,
  type CellAddress,
  type SerializedCellValue,
} from "../../../src/core/types";
import {
  cellAddressToKey,
  getCellReference,
  parseCellReference,
} from "../../../src/core/utils";

/**
 * Structured references to tables in other workbooks: `[Workbook]Table[Column]`
 * and its current-row, column-span and selector forms. Tables are workbook
 * resources, so a table in another workbook is named the way it is named in
 * its own workbook, with the workbook in front.
 */

const DATA_SHEET = "Data extraction";

function at(workbookName: string, sheetName: string, ref: string): CellAddress {
  return { workbookName, sheetName, ...parseCellReference(ref) };
}

/**
 * Add a batch workbook whose `Data extraction` sheet holds the table `Data`
 * with the columns Sample and Value.
 */
function addBatch(
  engine: FormulaEngine,
  workbookName: string,
  values: number[]
): void {
  const content = new Map<string, SerializedCellValue>([
    ["A1", "Sample"],
    ["B1", "Value"],
  ]);
  values.forEach((value, index) => {
    content.set(`A${index + 2}`, `S${index + 1}`);
    content.set(`B${index + 2}`, value);
  });
  engine.addWorkbook({
    workbookName,
    data: {
      sheets: [{ name: DATA_SHEET, content }],
      tables: [
        {
          name: "Data",
          sheetName: DATA_SHEET,
          start: "A1",
          numRows: { type: "number", value: values.length },
          numCols: 2,
        },
      ],
    },
  });
}

function addCollation(
  engine: FormulaEngine,
  content: Record<string, SerializedCellValue>
): void {
  engine.addWorkbook({
    workbookName: "Collation",
    data: {
      sheets: [{ name: "Summary", content: new Map(Object.entries(content)) }],
    },
  });
}

const summary = (ref: string) => at("Collation", "Summary", ref);

function value(engine: FormulaEngine, address: CellAddress): SerializedCellValue {
  return engine.getCellValue(address);
}

/** The stored content of a cell, such as its formula text. */
function formula(engine: FormulaEngine, address: CellAddress) {
  return engine
    .getSheetSerialized({
      workbookName: address.workbookName,
      sheetName: address.sheetName,
    })
    .get(getCellReference(address));
}

describe("cross-workbook structured references", () => {
  test("evaluate the references from the issue", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({
      workbookName: "B b",
      data: {
        sheets: [
          {
            name: "iQue RDA",
            content: new Map<string, SerializedCellValue>([
              ["A1", "Name"],
              ["B1", "Val"],
              ["A2", "x"],
              ["B2", 5],
              ["A3", "y"],
              ["B3", 7],
            ]),
          },
        ],
        tables: [
          {
            name: "RDA_1",
            sheetName: "iQue RDA",
            start: "A1",
            numRows: { type: "number", value: 2 },
            numCols: 2,
          },
        ],
      },
    });
    addCollation(engine, {
      A1: "=[B b]'iQue RDA'!B2*2",
      A2: "=INDEX([B b]RDA_1[Val],2)",
      // Tables are workbook resources: without the workbook, RDA_1 names a
      // table in Collation, which has none.
      A3: "=INDEX(RDA_1[Val],1)",
    });

    expect(value(engine, summary("A1"))).toBe(10);
    expect(value(engine, summary("A2"))).toBe(7);
    expect(value(engine, summary("A3"))).toBe(FormulaError.REF);
  });

  test("support column, selector, span and current-row forms", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [3, 4, 5]);
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])",
      A2: "=[Batch 1]Data[[#Headers],[Value]]",
      A3: "=COUNTA([Batch 1]Data[#Data])",
      A4: "=COUNTA([Batch 1]Data[#All])",
      A5: "=COUNTA([Batch 1]Data[[Sample]:[Value]])",
      A6: "=ROWS([Batch 1]Data[[#Data],[Sample]:[Value]])",
      // The current row is the formula's own row, as in its own workbook.
      C2: "=[Batch 1]Data[@Value]*10",
      C3: "=[Batch 1]Data[@[Value]]*10",
      C4: "=COUNTA([Batch 1]Data[@[Sample]:[Value]])",
      // A whole column spills.
      E1: "=[Batch 1]Data[Value]",
    });

    expect(value(engine, summary("A1"))).toBe(12);
    expect(value(engine, summary("A2"))).toBe("Value");
    expect(value(engine, summary("A3"))).toBe(6);
    expect(value(engine, summary("A4"))).toBe(8);
    expect(value(engine, summary("A5"))).toBe(6);
    expect(value(engine, summary("A6"))).toBe(3);
    expect(value(engine, summary("C2"))).toBe(30);
    expect(value(engine, summary("C3"))).toBe(40);
    expect(value(engine, summary("C4"))).toBe(2);
    expect(value(engine, summary("E1"))).toBe(3);
    expect(value(engine, summary("E2"))).toBe(4);
    expect(value(engine, summary("E3"))).toBe(5);
  });

  test("collate batch workbooks that use the same table name", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addBatch(engine, "Batch 2", [10, 20]);
    for (const workbookName of ["Batch 1", "Batch 2"]) {
      engine.setCellContent(at(workbookName, DATA_SHEET, "D1"), "=SUM(Data[Value])");
    }
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])+SUM([Batch 2]Data[Value])",
      A2: "=XLOOKUP(\"S2\",[Batch 2]Data[Sample],[Batch 2]Data[Value])",
    });

    expect(value(engine, at("Batch 1", DATA_SHEET, "D1"))).toBe(6);
    expect(value(engine, at("Batch 2", DATA_SHEET, "D1"))).toBe(30);
    expect(value(engine, summary("A1"))).toBe(36);
    expect(value(engine, summary("A2"))).toBe(20);
  });

  test("recalculate when the source table's data or extent changes", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])",
      A2: "=ROWS([Batch 1]Data[Value])",
      // Row 3 is the table's second data row.
      A3: "=[Batch 1]Data[@Value]",
    });
    expect(value(engine, summary("A1"))).toBe(6);
    expect(value(engine, summary("A2"))).toBe(3);
    expect(value(engine, summary("A3"))).toBe(2);

    engine.setCellContent(at("Batch 1", DATA_SHEET, "B3"), 100);
    expect(value(engine, summary("A1"))).toBe(104);
    expect(value(engine, summary("A3"))).toBe(100);

    engine.setCellContent(at("Batch 1", DATA_SHEET, "B5"), 1000);
    engine.updateTable({
      workbookName: "Batch 1",
      tableName: "Data",
      numRows: { type: "number", value: 4 },
    });
    expect(value(engine, summary("A1"))).toBe(1104);
    expect(value(engine, summary("A2"))).toBe(4);
  });

  test("follow rows appended to an open-ended table", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2]);
    engine.updateTable({
      workbookName: "Batch 1",
      tableName: "Data",
      numRows: { type: "infinity", sign: "positive" },
    });
    addCollation(engine, { A1: "=SUM([Batch 1]Data[Value])" });
    expect(value(engine, summary("A1"))).toBe(3);

    engine.setCellContent(at("Batch 1", DATA_SHEET, "B10"), 40);
    expect(value(engine, summary("A1"))).toBe(43);
  });

  test("resolve when the workbook or table appears and fail when it goes away", () => {
    const engine = FormulaEngine.buildEmpty();
    addCollation(engine, {
      A1: "=INDEX([Batch 1]Data[Value],1)",
      A2: "=[Batch 1]Data[[#Headers],[Value]]",
    });
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);
    expect(value(engine, summary("A2"))).toBe(FormulaError.REF);

    addBatch(engine, "Batch 1", [7, 8]);
    expect(value(engine, summary("A1"))).toBe(7);
    expect(value(engine, summary("A2"))).toBe("Value");

    engine.removeTable({ workbookName: "Batch 1", tableName: "Data" });
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);

    engine.addTable({
      workbookName: "Batch 1",
      sheetName: DATA_SHEET,
      tableName: "Data",
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });
    expect(value(engine, summary("A1"))).toBe(7);

    engine.removeWorkbook("Batch 1");
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);
    expect(value(engine, summary("A2"))).toBe(FormulaError.REF);

    addBatch(engine, "Batch 1", [9]);
    expect(value(engine, summary("A1"))).toBe(9);
    expect(value(engine, summary("A2"))).toBe("Value");
  });

  test("a missing column is #REF!", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1]);
    addCollation(engine, { A1: "=INDEX([Batch 1]Data[Missing],1)" });
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);
  });

  test("removing the table's sheet removes the table until undone", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [4, 5]);
    engine.addSheet({ workbookName: "Batch 1", sheetName: "Notes" });
    addCollation(engine, { A1: "=INDEX([Batch 1]Data[Value],2)" });
    expect(value(engine, summary("A1"))).toBe(5);

    engine.removeSheet({ workbookName: "Batch 1", sheetName: DATA_SHEET });
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);

    engine.undo();
    expect(value(engine, summary("A1"))).toBe(5);
  });

  test("a cloned workbook serves references written before it existed", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    engine.setCellContent(
      at("Batch 1", DATA_SHEET, "D1"),
      "=SUM([Batch 1]Data[Value])"
    );
    addCollation(engine, {
      A1: "=INDEX([Batch 2]Data[Value],1)",
      A2: "=SUM([Batch 2]Data[Value])",
    });
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);

    engine.cloneWorkbook("Batch 1", "Batch 2");

    expect(value(engine, summary("A1"))).toBe(1);
    expect(value(engine, summary("A2"))).toBe(6);
    // The clone's reference to its source workbook now names the clone.
    expect(formula(engine, at("Batch 2", DATA_SHEET, "D1"))).toBe(
      "=SUM([Batch 2]Data[Value])"
    );
    expect(value(engine, at("Batch 2", DATA_SHEET, "D1"))).toBe(6);

    engine.setCellContent(at("Batch 2", DATA_SHEET, "B2"), 10);
    expect(value(engine, summary("A1"))).toBe(10);
    expect(value(engine, summary("A2"))).toBe(15);
    expect(value(engine, at("Batch 1", DATA_SHEET, "D1"))).toBe(6);

    engine.undo();
    engine.undo();
    expect(engine.hasWorkbook("Batch 2")).toBe(false);
    expect(value(engine, summary("A1"))).toBe(FormulaError.REF);
  });

  test("renaming the source workbook rewrites the references", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])",
      A3: "=[Batch 1]Data[@Value]",
    });
    expect(value(engine, summary("A1"))).toBe(6);

    engine.renameWorkbook({
      workbookName: "Batch 1",
      newWorkbookName: "Batch A",
    });
    expect(formula(engine, summary("A1"))).toBe("=SUM([Batch A]Data[Value])");
    expect(formula(engine, summary("A3"))).toBe("=[Batch A]Data[@Value]");
    expect(value(engine, summary("A1"))).toBe(6);
    expect(value(engine, summary("A3"))).toBe(2);

    engine.setCellContent(at("Batch A", DATA_SHEET, "B2"), 10);
    expect(value(engine, summary("A1"))).toBe(15);

    engine.undo();
    engine.undo();
    expect(formula(engine, summary("A1"))).toBe("=SUM([Batch 1]Data[Value])");
    expect(value(engine, summary("A1"))).toBe(6);
  });

  test("renaming a table only rewrites references to that workbook's table", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addBatch(engine, "Batch 2", [10, 20]);
    for (const workbookName of ["Batch 1", "Batch 2"]) {
      engine.setCellContent(at(workbookName, DATA_SHEET, "D1"), "=SUM(Data[Value])");
    }
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])+SUM([Batch 2]Data[Value])",
      A3: "=[Batch 1]Data[@Value]",
    });

    engine.renameTable("Batch 1", { oldName: "Data", newName: "Results" });

    expect(formula(engine, summary("A1"))).toBe(
      "=SUM([Batch 1]Results[Value])+SUM([Batch 2]Data[Value])"
    );
    expect(formula(engine, summary("A3"))).toBe("=[Batch 1]Results[@Value]");
    expect(formula(engine, at("Batch 1", DATA_SHEET, "D1"))).toBe(
      "=SUM(Results[Value])"
    );
    // Batch 2's unqualified reference names its own Data table.
    expect(formula(engine, at("Batch 2", DATA_SHEET, "D1"))).toBe(
      "=SUM(Data[Value])"
    );
    expect(value(engine, summary("A1"))).toBe(36);
    expect(value(engine, summary("A3"))).toBe(2);
    expect(value(engine, at("Batch 1", DATA_SHEET, "D1"))).toBe(6);
    expect(value(engine, at("Batch 2", DATA_SHEET, "D1"))).toBe(30);

    engine.undo();
    expect(formula(engine, summary("A1"))).toBe(
      "=SUM([Batch 1]Data[Value])+SUM([Batch 2]Data[Value])"
    );
    expect(value(engine, summary("A1"))).toBe(36);

    engine.redo();
    expect(formula(engine, summary("A1"))).toBe(
      "=SUM([Batch 1]Results[Value])+SUM([Batch 2]Data[Value])"
    );
    expect(value(engine, summary("A1"))).toBe(36);
  });

  test("renaming a column header only rewrites references to that table", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addBatch(engine, "Batch 2", [10, 20]);
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])+SUM([Batch 2]Data[Value])",
      A3: "=[Batch 1]Data[@Value]",
      A4: "=COUNTA([Batch 1]Data[[Sample]:[Value]])",
    });

    engine.setCellContent(at("Batch 1", DATA_SHEET, "B1"), "Amount");

    expect(formula(engine, summary("A1"))).toBe(
      "=SUM([Batch 1]Data[Amount])+SUM([Batch 2]Data[Value])"
    );
    expect(formula(engine, summary("A3"))).toBe("=[Batch 1]Data[@Amount]");
    expect(formula(engine, summary("A4"))).toBe(
      "=COUNTA([Batch 1]Data[[Sample]:[Amount]])"
    );
    expect(value(engine, summary("A1"))).toBe(36);
    expect(value(engine, summary("A3"))).toBe(2);
    expect(value(engine, summary("A4"))).toBe(6);

    engine.undo();
    expect(formula(engine, summary("A1"))).toBe(
      "=SUM([Batch 1]Data[Value])+SUM([Batch 2]Data[Value])"
    );
    expect(formula(engine, summary("A3"))).toBe("=[Batch 1]Data[@Value]");
    expect(value(engine, summary("A1"))).toBe(36);
    expect(value(engine, summary("A3"))).toBe(2);
  });

  test("renaming the source sheet keeps references working", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addBatch(engine, "Batch 2", [10, 20]);
    engine.addSheet({ workbookName: "Batch 2", sheetName: "Notes" });
    engine.setCellContent(at("Batch 2", "Notes", "A1"), "='Data extraction'!B2");
    addCollation(engine, {
      A1: "=SUM([Batch 1]Data[Value])",
      A2: "=[Batch 1]'Data extraction'!B2",
      A3: "=[Batch 2]'Data extraction'!B2",
    });

    engine.renameSheet({
      workbookName: "Batch 1",
      sheetName: DATA_SHEET,
      newSheetName: "Extraction",
    });

    expect(formula(engine, summary("A1"))).toBe("=SUM([Batch 1]Data[Value])");
    expect(formula(engine, summary("A2"))).toBe("=[Batch 1]Extraction!B2");
    // Batch 2 has a sheet of the same name; references to it are unchanged.
    expect(formula(engine, summary("A3"))).toBe("=[Batch 2]'Data extraction'!B2");
    expect(formula(engine, at("Batch 2", "Notes", "A1"))).toBe(
      "='Data extraction'!B2"
    );
    expect(value(engine, summary("A1"))).toBe(6);
    expect(value(engine, summary("A2"))).toBe(1);
    expect(value(engine, summary("A3"))).toBe(10);
    expect(value(engine, at("Batch 2", "Notes", "A1"))).toBe(10);
  });

  test("named expressions can refer to tables in other workbooks", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addCollation(engine, { A1: "=BatchTotal" });
    engine.addNamedExpression({
      workbookName: "Collation",
      expressionName: "BatchTotal",
      expression: "SUM([Batch 1]Data[Value])",
    });
    expect(value(engine, summary("A1"))).toBe(6);

    engine.renameTable("Batch 1", { oldName: "Data", newName: "Results" });
    expect(
      engine._namedExpressionManager.getNamedExpression({
        name: "BatchTotal",
        scope: { type: "workbook", workbookName: "Collation" },
      })?.expression
    ).toBe("SUM([Batch 1]Results[Value])");
    expect(value(engine, summary("A1"))).toBe(6);

    engine.setCellContent(at("Batch 1", DATA_SHEET, "B2"), 11);
    expect(value(engine, summary("A1"))).toBe(16);
  });

  test("survive warm and compact serialization", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addCollation(engine, { A1: "=SUM([Batch 1]Data[Value])" });
    expect(value(engine, summary("A1"))).toBe(6);

    const warnings: unknown[][] = [];
    const originalWarn = console.warn;
    console.warn = (...args: unknown[]) => {
      warnings.push(args);
    };
    try {
      for (const includeEvaluationState of [true, false]) {
        const restored = FormulaEngine.buildEmpty();
        restored.resetToSerializedEngine(
          engine.serializeEngine({ includeEvaluationState })
        );
        if (includeEvaluationState) {
          // The warm snapshot restored the evaluated formula, not a cold engine.
          expect(
            restored._dependencyManager.getCellValueNode(
              cellAddressToKey(summary("A1"))
            ).resolved
          ).toBe(true);
        }
        expect(formula(restored, summary("A1"))).toBe(
          "=SUM([Batch 1]Data[Value])"
        );
        expect(value(restored, summary("A1"))).toBe(6);

        restored.setCellContent(at("Batch 1", DATA_SHEET, "B2"), 5);
        expect(value(restored, summary("A1"))).toBe(10);
      }
    } finally {
      console.warn = originalWarn;
    }
    expect(warnings).toEqual([]);
  });

  test("store formulas in canonical form", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addCollation(engine, {});
    engine.setCellContent(
      summary("A1"),
      "=SUM( [Batch 1]Data[[Sample]:[Value]] ) + [Batch 1]Data[@[Value]]"
    );
    expect(formula(engine, summary("A1"))).toBe(
      "=SUM([Batch 1]Data[[Sample]:[Value]])+[Batch 1]Data[@Value]"
    );
  });

  test("the former [Workbook]!Table[Column] spelling is a syntax error", () => {
    const engine = FormulaEngine.buildEmpty();
    addBatch(engine, "Batch 1", [1, 2, 3]);
    addCollation(engine, { A1: "=SUM([Batch 1]!Data[Value])" });
    expect(value(engine, summary("A1"))).toBe(FormulaError.ERROR);
  });
});
