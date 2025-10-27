import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";
import { visualizeSpreadsheet } from "../../../core/utils/spreadsheet-visualizer";

const sheetName = "TestSheet";
const workbookName = "TestWorkbook";
const sheetAddress = { workbookName, sheetName };
let engine: FormulaEngine;

const cell = (ref: string, debug?: boolean) =>
  engine.getCellValue(
    { sheetName, workbookName, ...parseCellReference(ref) },
    debug
  );

const setCellContent = (ref: string, content: string) => {
  engine.setCellContent(
    { sheetName, workbookName, ...parseCellReference(ref) },
    content
  );
};

const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

engine = FormulaEngine.buildEmpty();
engine.addWorkbook(workbookName);
engine.addSheet({ workbookName, sheetName });

engine.setSheetContent(
  sheetAddress,
  new Map<string, SerializedCellValue>([
    ["A1", "=B10000000"],
    ["B1", "=SEQUENCE(10000000)"],
  ])
);
// console.profile("evaluate")
console.log(cell("A1") === 10000000);
// console.profileEnd("evaluate");
