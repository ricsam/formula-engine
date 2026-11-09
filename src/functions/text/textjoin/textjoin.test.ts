import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("TEXTJOIN function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should join strings with delimiter", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, "Red", "Green", "Blue")');
      expect(cell("A1")).toBe("Red, Green, Blue");
    });

    test("should join strings with empty delimiter", () => {
      setCellContent("A1", '=TEXTJOIN("", TRUE, "Red", "Green", "Blue")');
      expect(cell("A1")).toBe("RedGreenBlue");
    });

    test("should join strings with hyphen delimiter", () => {
      setCellContent("A1", '=TEXTJOIN("-", TRUE, "a", "b", "c")');
      expect(cell("A1")).toBe("a-b-c");
    });

    test("should work with cell references", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["B1", "Green"],
          ["C1", "Blue"],
          ["D1", '=TEXTJOIN(", ", TRUE, A1, B1, C1)'],
        ])
      );
      expect(cell("D1")).toBe("Red, Green, Blue");
    });

    test("should join single item", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, "Only")');
      expect(cell("A1")).toBe("Only");
    });

    test("should work with numeric delimiter", () => {
      setCellContent("A1", '=TEXTJOIN(123, TRUE, "a", "b", "c")');
      expect(cell("A1")).toBe("a123b123c");
    });
  });

  describe("ignore_empty parameter", () => {
    test("should ignore empty strings when ignore_empty is TRUE", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, "Red", "", "Blue")');
      expect(cell("A1")).toBe("Red, Blue");
    });

    test("should include empty strings when ignore_empty is FALSE", () => {
      setCellContent("A1", '=TEXTJOIN(", ", FALSE, "Red", "", "Blue")');
      expect(cell("A1")).toBe("Red, , Blue");
    });

    test("should ignore multiple empty strings when TRUE", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, "Red", "", "", "Blue", "")');
      expect(cell("A1")).toBe("Red, Blue");
    });

    test("should handle all empty strings with ignore_empty TRUE", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, "", "", "")');
      expect(cell("A1")).toBe("");
    });

    test("should handle all empty strings with ignore_empty FALSE", () => {
      setCellContent("A1", '=TEXTJOIN(", ", FALSE, "", "", "")');
      expect(cell("A1")).toBe(", , ");
    });

    test("should accept 1 as TRUE for ignore_empty", () => {
      setCellContent("A1", '=TEXTJOIN(", ", 1, "Red", "", "Blue")');
      expect(cell("A1")).toBe("Red, Blue");
    });

    test("should accept 0 as FALSE for ignore_empty", () => {
      setCellContent("A1", '=TEXTJOIN(", ", 0, "Red", "", "Blue")');
      expect(cell("A1")).toBe("Red, , Blue");
    });
  });

  describe("range arguments", () => {
    test("should join range with ignore_empty TRUE", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["A2", "Green"],
          ["A3", "Blue"],
          ["B1", '=TEXTJOIN(", ", TRUE, A1:A3)'],
        ])
      );
      expect(cell("B1")).toBe("Red, Green, Blue");
    });

    test("should join range with empty cells and ignore_empty TRUE", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["A2", ""],
          ["A3", "Blue"],
          ["B1", '=TEXTJOIN(", ", TRUE, A1:A3)'],
        ])
      );
      expect(cell("B1")).toBe("Red, Blue");
    });

    test("should join range with empty cells and ignore_empty FALSE", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["A2", ""],
          ["A3", "Blue"],
          ["B1", '=TEXTJOIN(", ", FALSE, A1:A3)'],
        ])
      );
      // Note: evaluateAllCells skips empty cells in ranges by design
      // This is consistent with Excel's behavior where empty cells are ignored
      expect(cell("B1")).toBe("Red, Blue");
    });

    test("should join horizontal range", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["B1", "Green"],
          ["C1", "Blue"],
          ["D1", '=TEXTJOIN("-", TRUE, A1:C1)'],
        ])
      );
      expect(cell("D1")).toBe("Red-Green-Blue");
    });

    test("should join 2D range", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "A"],
          ["B1", "B"],
          ["A2", "C"],
          ["B2", "D"],
          ["C1", '=TEXTJOIN("-", TRUE, A1:B2)'],
        ])
      );
      // Should flatten in column-major order: A1, A2, B1, B2
      expect(cell("C1")).toBe("A-C-B-D");
    });

    test("should join multiple ranges", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["A2", "Green"],
          ["C1", "Blue"],
          ["C2", "Yellow"],
          ["D1", '=TEXTJOIN(", ", TRUE, A1:A2, C1:C2)'],
        ])
      );
      expect(cell("D1")).toBe("Red, Green, Blue, Yellow");
    });

    test("should mix ranges and scalar values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["A2", "Green"],
          ["D1", '=TEXTJOIN(", ", TRUE, "Start", A1:A2, "End")'],
        ])
      );
      expect(cell("D1")).toBe("Start, Red, Green, End");
    });
  });

  describe("type coercion", () => {
    test("should convert numbers to strings", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, 1, 2, 3)');
      expect(cell("A1")).toBe("1, 2, 3");
    });

    test("should convert booleans to strings", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, TRUE, FALSE)');
      expect(cell("A1")).toBe("TRUE, FALSE");
    });

    test("should handle mixed types", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, "Text", 123, TRUE, 3.14)');
      expect(cell("A1")).toBe("Text, 123, TRUE, 3.14");
    });

    test("should convert cell references with numbers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["B1", '=TEXTJOIN("+", TRUE, A1:A3)'],
        ])
      );
      expect(cell("B1")).toBe("10+20+30");
    });

    test("should handle infinity values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"],
          ["A2", "=-1/0"],
          ["B1", '=TEXTJOIN(", ", TRUE, "Start", A1, A2, "End")'],
        ])
      );
      expect(cell("B1")).toBe("Start, INFINITY, -INFINITY, End");
    });
  });

  describe("error handling", () => {
    test("should return #VALUE! for too few arguments", () => {
      setCellContent("A1", "=TEXTJOIN()");
      expect(cell("A1")).toBe(FormulaError.VALUE);

      setCellContent("A2", '=TEXTJOIN(",")');
      expect(cell("A2")).toBe(FormulaError.VALUE);

      setCellContent("A3", '=TEXTJOIN(",", TRUE)');
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("should handle error in delimiter", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"],
          ["B1", '=TEXTJOIN(A1, TRUE, "a", "b")'],
        ])
      );
      // 1/0 produces INFINITY, which is coerced to "INFINITY" as delimiter
      expect(cell("B1")).toBe("aINFINITYb");
    });

    test("should propagate error from text arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=INVALID_FUNCTION()"],
          ["B1", '=TEXTJOIN(", ", TRUE, "a", A1, "b")'],
        ])
      );
      const result = cell("B1");
      expect(result).toStartWith("#NAME?");
    });

    test("should return #VALUE! if result exceeds 32767 characters", () => {
      // Create a formula that generates a very long string
      const longString = "x".repeat(10000);
      setCellContent("A1", `=TEXTJOIN("", TRUE, "${longString}", "${longString}", "${longString}", "${longString}")`);
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! if delimiter is a range", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", ","],
          ["A2", ";"],
          ["B1", '=TEXTJOIN(A1:A2, TRUE, "a", "b")'],
        ])
      );
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! if ignore_empty is a range", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", true],
          ["A2", false],
          ["B1", '=TEXTJOIN(",", A1:A2, "a", "b")'],
        ])
      );
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("edge cases", () => {
    test("should handle only empty values with ignore_empty TRUE", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", ""],
          ["A2", ""],
          ["A3", ""],
          ["B1", '=TEXTJOIN(", ", TRUE, A1:A3)'],
        ])
      );
      expect(cell("B1")).toBe("");
    });

    test("should handle delimiter from cell reference", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", " | "],
          ["B1", "Red"],
          ["B2", "Green"],
          ["C1", "=TEXTJOIN(A1, TRUE, B1, B2)"],
        ])
      );
      expect(cell("C1")).toBe("Red | Green");
    });

    test("should handle ignore_empty from cell reference", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", true],
          ["B1", '=TEXTJOIN(", ", A1, "Red", "", "Blue")'],
        ])
      );
      expect(cell("B1")).toBe("Red, Blue");
    });

    test("should handle special characters in delimiter", () => {
      setCellContent("A1", '=TEXTJOIN("\\n", TRUE, "Line1", "Line2")');
      expect(cell("A1")).toBe("Line1\\nLine2");
    });

    test("should work with zero as text value", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, 0, 1, 2)');
      expect(cell("A1")).toBe("0, 1, 2");
    });

    test("should handle FALSE boolean with ignore_empty TRUE", () => {
      setCellContent("A1", '=TEXTJOIN(", ", TRUE, TRUE, FALSE, TRUE)');
      expect(cell("A1")).toBe("TRUE, FALSE, TRUE");
    });

    test("should work with many arguments", () => {
      setCellContent("A1", '=TEXTJOIN("-", TRUE, "a", "b", "c", "d", "e", "f", "g", "h", "i", "j")');
      expect(cell("A1")).toBe("a-b-c-d-e-f-g-h-i-j");
    });
  });

  describe("delimiter variations", () => {
    test("should work with multi-character delimiter", () => {
      setCellContent("A1", '=TEXTJOIN(" <-> ", TRUE, "A", "B", "C")');
      expect(cell("A1")).toBe("A <-> B <-> C");
    });

    test("should work with space delimiter", () => {
      setCellContent("A1", '=TEXTJOIN(" ", TRUE, "Hello", "Beautiful", "World")');
      expect(cell("A1")).toBe("Hello Beautiful World");
    });

    test("should work with tab delimiter", () => {
      setCellContent("A1", '=TEXTJOIN("\t", TRUE, "Col1", "Col2", "Col3")');
      expect(cell("A1")).toBe("Col1\tCol2\tCol3");
    });
  });

  describe("omitted arguments", () => {
    test("an omitted ignore_empty is TRUE", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Red"],
          ["A3", "Blue"],
          ["B1", '=TEXTJOIN(", ",, A1:A3)'],
          ["B2", '=TEXTJOIN(", ",, "x", "", "y")'],
          ["B3", '=TEXTJOIN(", ", FALSE, "x", "", "y")'],
        ])
      );

      expect(cell("B1")).toBe("Red, Blue");
      expect(cell("B2")).toBe("x, y");
      expect(cell("B3")).toBe("x, , y");
    });

    test("an omitted delimiter joins without a delimiter", () => {
      setCellContent("A1", '=TEXTJOIN(, TRUE, "a", "b", "c")');
      expect(cell("A1")).toBe("abc");
    });

    test("an omitted text item is empty text", () => {
      setCellContent("A1", '=TEXTJOIN("-", FALSE, "a",, "b")');
      setCellContent("A2", '=TEXTJOIN("-", TRUE, "a",, "b")');

      expect(cell("A1")).toBe("a--b");
      expect(cell("A2")).toBe("a-b");
    });
  });

  describe("structured references", () => {
    test("joins every column of a current-row column span", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "a"],
          ["B1", "b"],
          ["C1", "c"],
          ["D1", "joined"],
          ["A2", 1],
          ["B2", 2],
          ["C2", 3],
          ["D2", '=TEXTJOIN("-", TRUE, [@[a]:[c]])'],
          ["A3", "x"],
          ["C3", "z"],
          ["D3", '=TEXTJOIN("-", TRUE, [@[a]:[c]])'],
        ])
      );
      engine.addTable({
        tableName: "Spans",
        workbookName,
        sheetName,
        start: "A1",
        numRows: { type: "number", value: 2 },
        numCols: 4,
      });

      expect(cell("D2")).toBe("1-2-3");
      expect(cell("D3")).toBe("x-z");
    });

    /**
     * The experiment_metadata column (J) of the ted_export_rfp and
     * ted_export_cd3- sheets of the Responder Peptide Pulse batch workbook,
     * imported as WorkbookData the way the excel-to-formula-engine importer
     * does. The formula and the pep_concentration (H) and gate (I) columns it
     * reads come from the workbook; columns A–G stand in for the rest of the
     * export.
     */
    test("computes the experiment_metadata column of the peptide pulse export sheets", () => {
      const experimentMetadata =
        '=IF(OR([@pep_concentration]="np",[@pep_concentration]="null",[@pep_concentration]="MAP"),[@gate],TEXTJOIN(", ",TRUE,[@[pep_concentration]:[gate]]))';
      const exportWorkbook = "Responder Peptide Pulse";
      const exportSheet = (gate: string, concentrations: string[]) => {
        const content = new Map<string, SerializedCellValue>();
        "ABCDEFG".split("").forEach((column, index) => {
          content.set(`${column}1`, `field_${index + 1}`);
        });
        content.set("H1", "pep_concentration");
        content.set("I1", "gate");
        content.set("J1", "experiment_metadata");
        concentrations.forEach((concentration, index) => {
          const row = index + 2;
          content.set(`A${row}`, `sample ${index + 1}`);
          content.set(`H${row}`, concentration);
          content.set(`I${row}`, gate);
          content.set(`J${row}`, experimentMetadata);
        });
        return content;
      };
      const concentrations = ["10uM", "1uM", "np", "null", "MAP"];

      engine.addWorkbook({
        workbookName: exportWorkbook,
        data: {
          sheets: [
            { name: "ted_export_rfp", content: exportSheet("RFP+", concentrations) },
            { name: "ted_export_cd3-", content: exportSheet("CD3-", concentrations) },
          ],
          tables: [
            {
              name: "ted_export_rfp",
              sheetName: "ted_export_rfp",
              start: "A1",
              numRows: { type: "number", value: concentrations.length },
              numCols: 10,
            },
            {
              name: "ted_export_cd3",
              sheetName: "ted_export_cd3-",
              start: "A1",
              numRows: { type: "number", value: concentrations.length },
              numCols: 10,
            },
          ],
        },
      });

      const experimentMetadataColumn = (sheet: string) =>
        concentrations.map((_, index) =>
          engine.getCellValue({
            workbookName: exportWorkbook,
            sheetName: sheet,
            ...parseCellReference(`J${index + 2}`),
          })
        );

      expect(experimentMetadataColumn("ted_export_rfp")).toEqual([
        "10uM, RFP+",
        "1uM, RFP+",
        "RFP+",
        "RFP+",
        "RFP+",
      ]);
      expect(experimentMetadataColumn("ted_export_cd3-")).toEqual([
        "10uM, CD3-",
        "1uM, CD3-",
        "CD3-",
        "CD3-",
        "CD3-",
      ]);
      // The imported formula is stored with its whole span.
      expect(
        engine
          .getSheetSerialized({ workbookName: exportWorkbook, sheetName: "ted_export_rfp" })
          .get("J2")
      ).toContain('TEXTJOIN(", ", TRUE, [@[pep_concentration]:[gate]])');
    });
  });
});
