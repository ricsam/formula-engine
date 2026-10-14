import { expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { parseCellReference } from "../../../src/core/utils";

test("single-cell ranges retain explicit workbook context", () => {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName: "Inputs", data: { sheets: [{ name: "Sheet", content: { A1: 42 } }] } });
  engine.addWorkbook({ workbookName: "Output", data: { sheets: [{ name: "Sheet", content: { A1: 7, B1: "=SUM([Inputs]Sheet!A1:A1)" } }] } });
  expect(engine.getCellValue({ workbookName: "Output", sheetName: "Sheet", ...parseCellReference("B1") })).toBe(42);
});

for (const reverse of [false, true]) {
  test(`range aggregate caches respect workbook and sheet context (reverse=${reverse})`, () => {
    const engine = FormulaEngine.buildEmpty();
    for (const workbookName of ["Original", "Imported"]) {
      engine.addWorkbook({ workbookName, data: { sheets: ["First", "Second"].map((name, i) => ({
        name, content: { A1: (workbookName === "Original" ? 10 : 100) + i, A2: 5, B1: "=SUM(A1:A2)" },
      })) } });
    }
    const addresses = ["Original", "Imported"].flatMap(workbookName => ["First", "Second"].map(sheetName => ({
      workbookName, sheetName, ...parseCellReference("B1"),
    })));
    if (reverse) addresses.reverse();
    for (const address of addresses) {
      expect(engine.getCellValue(address)).toBe((address.workbookName === "Original" ? 15 : 105) + (address.sheetName === "Second" ? 1 : 0));
    }
  });
}

for (const warmDependencies of [false, true]) {
  test(`nested range totals are independent of read order (warm=${warmDependencies})`, () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: "Assumptions", data: { sheets: [{
      name: "Inputs", content: { B4: 45, B5: 2400, B6: 0.1 },
    }] } });
    engine.addWorkbook({ workbookName: "Launch budget", data: { sheets: [{
      name: "Overview", content: {
        B5: 80, C5: "=[Assumptions]Inputs!B4", D5: "=B5*C5",
        B6: 1, C6: "=[Assumptions]Inputs!B5", D6: "=B6*C6",
        B7: 1, C7: 850, D7: "=B7*C7",
        D9: "=SUM(D5:D7)", D10: "=D9*[Assumptions]Inputs!B6",
        D12: "=SUM(D9:D10)",
      },
    }] } });
    const address = (ref: string) => ({
      workbookName: "Launch budget", sheetName: "Overview", ...parseCellReference(ref),
    });
    const read = (ref: string) => engine.getCellValue(address(ref));
    if (warmDependencies) {
      expect(read("D9")).toBe(6850);
      expect(read("D10")).toBe(685);
    }
    expect(read("D12")).toBe(7535);
    expect(read("D10")).toBe(685);
    expect(read("D12")).toBe(7535);
    engine.setCellContent({ workbookName: "Assumptions", sheetName: "Inputs", ...parseCellReference("B6") }, 0.2);
    expect(read("D12")).toBe(8220);
    engine.setCellContent({ workbookName: "Assumptions", sheetName: "Inputs", ...parseCellReference("B4") }, 50);
    expect(read("D12")).toBe(8700);
  });
}
