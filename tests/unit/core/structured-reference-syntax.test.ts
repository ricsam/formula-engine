import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { CellAddress, SerializedCellValue } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

/**
 * Which structured-reference spellings the engine actually accepts.
 *
 * Excel has several ways of writing the same reference, and an importer reading
 * a .xlsx meets the *stored* spelling rather than the one the author typed. This
 * pins down which spellings evaluate, so that an importer knows what it has to
 * rewrite and what it can pass through untouched.
 */

const workbookName = "Book";
const sheetName = "Sheet1";

/**
 * A sheet holding a three-column table `TTRinput` over A1:C3, with `formula`
 * placed in C2.
 *
 * The formula cell is *inside* the table — a calculated column, which is what
 * Excel creates when you type a formula beside table data. That matters: an
 * unqualified `[@Column]` names no table, so it can only resolve from the cell
 * it sits in, and a formula outside the table has nothing to resolve against.
 */
function buildEngine(formula: string): FormulaEngine {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({
    workbookName,
    data: {
      sheets: [
        {
          name: sheetName,
          content: new Map<string, SerializedCellValue>([
            ["A1", "Payload"],
            ["B1", "Amount"],
            ["C1", "Result"],
            ["A2", "alpha,beta"],
            ["B2", 10],
            ["A3", "gamma,delta"],
            ["B3", 20],
            ["C2", formula],
          ]),
        },
      ],
      tables: [
        {
          name: "TTRinput",
          sheetName,
          start: "A1",
          numRows: { type: "number", value: 2 },
          numCols: 3,
        },
      ],
    },
  });
  return engine;
}

function valueAt(engine: FormulaEngine, ref: string): SerializedCellValue {
  const { rowIndex, colIndex } = parseCellReference(ref);
  const address: CellAddress = { workbookName, sheetName, rowIndex, colIndex };
  return engine.getCellValue(address);
}

/** Evaluate `formula` in C2 and report what the engine made of it. */
function evaluate(formula: string): SerializedCellValue {
  return valueAt(buildEngine(formula), "C2");
}

describe("current-row structured references", () => {
  test("[@Column] — the spelling Excel shows in its formula bar", () => {
    // "alpha,beta" up to the comma.
    expect(evaluate("=LEFT([@Payload], FIND(\",\", [@Payload])-1)")).toBe("alpha");
  });

  test("Table[@Column] — qualified with the table name", () => {
    expect(evaluate("=LEFT(TTRinput[@Payload], FIND(\",\", TTRinput[@Payload])-1)")).toBe(
      "alpha"
    );
  });

  test("Table[@[Column]] — qualified, with the column bracketed", () => {
    expect(evaluate("=TTRinput[@[Amount]]")).toBe(10);
  });

  test("[@[Column]] — unqualified, with the column bracketed", () => {
    expect(evaluate("=[@[Amount]]")).toBe(10);
  });

  /**
   * The `[#This Row]` selector is NOT supported, in either spelling.
   *
   * This matters out of proportion to how obscure it looks, because it is the
   * spelling a .xlsx actually *stores*: Excel writes the fully-qualified
   * `Table[[#This Row],[Column]]` into the file and only ever displays `[@Column]`
   * in its formula bar. So it is the form an importer meets for every calculated
   * column, while being the one form the engine cannot read.
   *
   * These are characterization tests: they pin the current behaviour rather than
   * bless it. Teaching the parser this selector should make them fail — at which
   * point they become the assertions above.
   */
  test("Table[[#This Row],[Column]] does not parse", () => {
    expect(evaluate("=TTRinput[[#This Row],[Amount]]")).toBe("#ERROR!");
  });

  test("Table[[#This Row],[Column]] does not parse inside a function either", () => {
    expect(
      evaluate(
        '=LEFT(TTRinput[[#This Row],[Payload]],FIND(",",TTRinput[[#This Row],[Payload]])-1)'
      )
    ).toBe("#ERROR!");
  });

  /**
   * The unspaced spelling gets further — the lexer knows the word — but the
   * evaluator has no current-row meaning to give it, so it resolves to nothing.
   */
  test("Table[[#ThisRow],[Column]] lexes but resolves to #REF!", () => {
    expect(evaluate("=TTRinput[[#ThisRow],[Amount]]")).toBe("#REF!");
  });
});

describe("whole-column structured references", () => {
  test("Table[Column] reads the column's data", () => {
    expect(evaluate("=SUM(TTRinput[Amount])")).toBe(30);
  });

  test("Table[[Column]] with the column bracketed", () => {
    expect(evaluate("=SUM(TTRinput[[Amount]])")).toBe(30);
  });
});

describe("selector structured references", () => {
  test("Table[#Data] covers the data rows", () => {
    expect(evaluate("=SUM(TTRinput[#Data])")).toBe(30);
  });

  test("Table[[#Data],[Column]] narrows a selector to one column", () => {
    expect(evaluate("=SUM(TTRinput[[#Data],[Amount]])")).toBe(30);
  });

  test("Table[[#Headers],[Column]] reads the header cell", () => {
    expect(evaluate("=TTRinput[[#Headers],[Amount]]")).toBe("Amount");
  });
});
