/**
 * Excel's approximate-match lookup: the binary search behind MATCH with
 * match_type 1 or -1, which VLOOKUP, HLOOKUP and LOOKUP share in Excel.
 *
 * Excel never checks that the lookup array is sorted. It halves the searched
 * part of the array and returns where that ends, so sorted data gives the
 * largest value less than or equal to the lookup value (the smallest greater
 * or equal one for descending data), and unsorted data gives the position the
 * same probes lead to. Workbooks rely on both: `MATCH(x, ids)` without the
 * `,0` usually finds an exact hit, and `MATCH(9.99E+307, A:A)` finds the last
 * number in a column.
 *
 * The rules below reproduce published Excel results on unsorted data,
 * duplicates, blank cells, errors and mixed types (Excel's cached values in
 * Apache POI's lookup test workbook among them), and follow Apache POI's
 * `LookupUtils.performBinarySearch` where those results leave a choice open:
 *
 * - The search covers positions `low..high`, initially the whole array, and
 *   probes `floor((low + high) / 2)`.
 * - A probed cell that cannot be compared with the lookup value (a blank, an
 *   error or a value of another type) is stepped over: the probe moves forward
 *   to the first comparable cell up to `high`. Without one, `high` becomes the
 *   cell before the probe.
 * - A cell that sorts before the lookup value is the answer so far, and the
 *   search continues after it. A cell that sorts after it ends the search
 *   range: at the cell before the probe, or, when the probe stepped forward to
 *   the cell, before that cell unless it is the last one of the range.
 * - An equal cell ends the search. The result moves on while the next cells
 *   are equal too, so the last of adjacent duplicates is found; any other cell,
 *   a blank included, stops it.
 * - Without an equal cell the answer is the last cell that sorted before the
 *   lookup value, and none means `#N/A`.
 */

/** A cell of a lookup array, compared with the lookup value. */
export type ApproximateMatchCell = {
  /** 0-based position of the cell in the lookup array. */
  position: number;
  /**
   * The cell compared with the lookup value in ascending order: negative when
   * the cell is smaller, 0 when they are equal and positive when the cell is
   * larger. `null` when they cannot be compared, for an error or a value of
   * another type.
   */
  comparison: number | null;
};

/**
 * Returns the 0-based position Excel's approximate match finds, or
 * `undefined` for `#N/A`.
 *
 * @param cells the cells of the lookup array in ascending position order;
 *   positions without a cell are blank.
 * @param length the number of positions searched.
 * @param order "ascending" for match_type 1, "descending" for match_type -1,
 *   which searches as if every comparison were reversed.
 */
export function findApproximateMatch(
  cells: readonly ApproximateMatchCell[],
  length: number,
  order: "ascending" | "descending"
): number | undefined {
  const direction = order === "ascending" ? 1 : -1;

  // nextComparable[i] is the index of the first comparable cell at or after
  // cells[i], or cells.length when there is none.
  const nextComparable = new Array<number>(cells.length + 1);
  nextComparable[cells.length] = cells.length;
  for (let index = cells.length - 1; index >= 0; index--) {
    nextComparable[index] =
      cells[index]!.comparison === null ? nextComparable[index + 1]! : index;
  }

  /** The index of the first cell at `position` or after it. */
  const firstCellFrom = (position: number): number => {
    let low = 0;
    let high = cells.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (cells[middle]!.position < position) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    return low;
  };

  let low = 0;
  let high = length - 1;
  let found: number | undefined;

  while (low <= high) {
    const probe = Math.floor((low + high) / 2);
    const index = nextComparable[firstCellFrom(probe)]!;
    const cell = cells[index];

    if (cell === undefined || cell.position > high) {
      // Nothing from the probe to the end of the range can be compared.
      high = probe - 1;
      continue;
    }

    const comparison = direction * cell.comparison!;
    if (comparison === 0) {
      return lastOfEqualRun(cells, index, high);
    }

    if (comparison < 0) {
      found = cell.position;
      low = cell.position + 1;
    } else if (cell.position === probe || cell.position === high) {
      high = probe - 1;
    } else {
      high = cell.position - 1;
    }
  }

  return found;
}

/**
 * Returns the position of the last cell in the run of equal cells starting at
 * `cells[index]`, looking no further than position `high`.
 */
function lastOfEqualRun(
  cells: readonly ApproximateMatchCell[],
  index: number,
  high: number
): number {
  let last = index;
  for (let next = index + 1; next < cells.length; next++) {
    const cell = cells[next]!;
    if (
      cell.position !== cells[last]!.position + 1 ||
      cell.position > high ||
      cell.comparison !== 0
    ) {
      break;
    }
    last = next;
  }
  return cells[last]!.position;
}
