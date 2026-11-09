import { describe, expect, test } from "bun:test";
import {
  findApproximateMatch,
  type ApproximateMatchCell,
} from "./approximate-match";

/** A cell that cannot be compared with the lookup value, such as an error. */
const MISMATCH = Symbol("mismatch");

/** undefined is a blank cell. */
type Item = number | string | typeof MISMATCH | undefined;

function compare(item: number | string, lookup: number | string) {
  if (typeof item !== typeof lookup) return null;
  if (typeof item === "number") return item - (lookup as number);
  const left = item.toLowerCase();
  const right = (lookup as string).toLowerCase();
  return left < right ? -1 : left > right ? 1 : 0;
}

/** Returns the 1-based position found, like MATCH, or "#N/A". */
function search(
  items: Item[],
  lookup: number | string,
  order: "ascending" | "descending" = "ascending"
) {
  const cells: ApproximateMatchCell[] = [];
  items.forEach((item, position) => {
    if (item === undefined) return;
    cells.push({
      position,
      comparison: item === MISMATCH ? null : compare(item, lookup),
    });
  });
  const position = findApproximateMatch(cells, items.length, order);
  return position === undefined ? "#N/A" : position + 1;
}

const blanks = (count: number): Item[] => Array<Item>(count).fill(undefined);

describe("findApproximateMatch", () => {
  test("finds the largest value at most the lookup value in sorted data", () => {
    expect(search([10, 20, 30, 40], 25)).toBe(2);
    expect(search([10, 20, 30, 40], 40)).toBe(4);
    expect(search([10, 20, 30, 40], 45)).toBe(4);
    expect(search([10, 20, 30, 40], 5)).toBe("#N/A");
    expect(search([10, 20, 20, 30], 20)).toBe(3);
  });

  test("searches descending data with every comparison reversed", () => {
    expect(search([40, 30, 20, 10], 25, "descending")).toBe(2);
    expect(search([30, 20, 20, 10], 20, "descending")).toBe(3);
    expect(search([30, 20, 20, 10], 35, "descending")).toBe("#N/A");
    // Excel: MATCH(0, {3,2,2}, -1) is 3.
    expect(search([3, 2, 2], 0, "descending")).toBe(3);
  });

  test("probes the lower middle and returns where unsorted data leads", () => {
    // The issue: Excel returns 3 for MATCH("MAP-994", A1:A3).
    expect(search(["MAP-994", "MAP-10977", "MAP-320"], "MAP-994")).toBe(3);
    // Unsorted tables of Apache POI's lookup test workbook, with the results
    // Excel cached for VLOOKUP(..., TRUE).
    expect(search([33, 33, 33, 100, 33, 33, 33, 33, 33, 33], 33)).toBe(10);
    expect(search([33, 33, 33, 33, 100, 33, 33, 33, 33, 33], 33)).toBe(4);
    expect(search([32, 32, 32, 31, 100, 32, 32, 32, 32, 32], 33)).toBe(4);
    const table5 = [
      33, 33, 33, 33, 33, 33, 33, 33, 33, 33, 8, 33, 33, 6, 33, 33, 4, 33, 33,
      33, 33,
    ];
    expect(search(table5, 33)).toBe(16);
    expect(search(table5, 10)).toBe(11);
  });

  test("moves on from an exact hit while the next cells are equal", () => {
    // Excel's results from Microsoft Q&A (2023).
    expect(search([1, 1, 1, 1, 1, 0, 1], 1)).toBe(5);
    expect(
      search([1, 1, 1, 1, 1, 0, 1, -1, 1, 1, 1, 1, 1, 0, 1], 1)
    ).toBe(13);
    expect(search([1, 1, 1, 1, 1, 0, 1, 2, 1, 1, 1, 1, 1, 0, 1], 1)).toBe(5);
    // A blank or a cell of another type ends the run.
    expect(search([1, 1, 1, 1], 1)).toBe(4);
    expect(search([1, 1, undefined, 1], 1)).toBe(2);
    expect(search([1, 1, MISMATCH, 1], 1)).toBe(2);
  });

  test("steps forward over cells it cannot compare", () => {
    // Excel's results from the MrExcel forum (2016): MATCH("z", A1:A18) is 15
    // and MATCH("Ω", A1:A18) is 17 with blank cells in A8:A11.
    const column: Item[] = [
      "w", "x", "y", "z", "zz", "Ω", "ΩΩ", ...blanks(4),
      "w", "x", "y", "z", "zz", "Ω", "ΩΩ",
    ];
    expect(search(column, "z")).toBe(15);
    expect(search(column, "Ω")).toBe(17);

    // Excel: MATCH(9.99E+307, 1/(A1:A18="z")) finds the last 1 among errors.
    const ones: Item[] = Array<Item>(18).fill(MISMATCH);
    ones[3] = 1;
    ones[14] = 1;
    expect(search(ones, 9.99e307)).toBe(15);

    // Excel: MATCH("Ω", A:A) is 10 with only A1 "z" and A10 "a".
    const wholeColumn = blanks(1048576);
    wholeColumn[0] = "z";
    wholeColumn[9] = "a";
    expect(search(wholeColumn, "Ω")).toBe(10);
  });

  test("ends the range before a larger cell it stepped to, unless that cell ends the range", () => {
    // Apache POI's reading of Excel. The first probe, position 3, is blank and
    // steps forward to position 4, whose 3 is larger than 2. Position 4 is not
    // the last of the range, so the range ends at position 3 and the next
    // probe finds the 1 at position 2.
    expect(search([3, 1, undefined, 3, 1], 2)).toBe(2);
    // Here the probe steps forward to position 5, the last of the range, so
    // the range ends before the probe instead and the 1 is never probed.
    expect(search([3, 1, undefined, undefined, 3], 2)).toBe("#N/A");
  });

  test("finds nothing in an empty lookup array", () => {
    expect(search([], 1)).toBe("#N/A");
    expect(search(blanks(3), 1)).toBe("#N/A");
    expect(search([MISMATCH, "a"], 1)).toBe("#N/A");
  });
});
