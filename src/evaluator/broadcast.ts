import {
  FormulaError,
  type CellAddress,
  type CellInRangeResult,
  type EvaluateAllCellsResult,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpilledValuesEvaluationResult,
  type SpreadsheetRange,
} from "../core/types";
import { getResultShape, type ArrayShape } from "./array-shape";
import type { EvaluationContext } from "./evaluation-context";
import type { FormulaEvaluator } from "./formula-evaluator";

/**
 * Excel array broadcasting ("array expansion") for element-wise operations such
 * as the binary operators and IF with an array condition.
 *
 * Every operand is viewed as an array; a single value is 1x1. In each dimension
 * the result is as large as the largest operand. An operand of size 1 in a
 * dimension is repeated along it, and an operand larger than 1 but smaller than
 * the result yields `#N/A` in the cells beyond it:
 *
 * ```
 * SEQUENCE(3) * SEQUENCE(1, 4)   -> 3x4 multiplication table
 * {1;2;3} + {10;20}              -> {11; 22; #N/A}
 * ```
 *
 * Unbounded dimensions (open-ended ranges such as `A:A`) have size `Infinity`.
 */

/** One argument of an element-wise operation. */
export type BroadcastOperand =
  | {
      /** A single value, repeated over the whole result. */
      type: "scalar";
      /**
       * Usually a value. An error is a valid operand too (IF chooses it only in
       * the cells that take its branch). Return awaiting-evaluation results before
       * building operands: they are not values.
       */
      result: SingleEvaluationResult;
    }
  | {
      type: "array";
      result: SpilledValuesEvaluationResult;
      shape: ArrayShape;
    };

/**
 * Computes one cell of an element-wise result. `operandCell(i)` returns operand
 * `i`'s value for that cell and is only evaluated when called, so an operation
 * can read just the operands it needs (IF reads one branch per cell).
 */
export type BroadcastCombine = (
  operandCell: (operandIndex: number) => SingleEvaluationResult
) => SingleEvaluationResult;

/** A cell offset relative to the start of an array. */
type Offset = { x: number; y: number };

/**
 * An inclusive block of cells relative to the start of an array. `x1`/`y1` are
 * `Infinity` for unbounded dimensions.
 */
type Region = { x0: number; x1: number; y0: number; y1: number };

const BLANK: SingleEvaluationResult = {
  type: "value",
  result: { type: "string", value: "" },
};

/**
 * Wraps an evaluated argument as a broadcast operand. A spilled result gets the
 * shape of the area it occupies when spilled from `origin`.
 */
export function toBroadcastOperand(
  result: FunctionEvaluationResult,
  origin: CellAddress
): BroadcastOperand {
  if (result.type === "spilled-values") {
    return {
      type: "array",
      result,
      shape: getResultShape(result, origin),
    };
  }
  return { type: "scalar", result };
}

/** The shape of an operand; a single value is 1x1. */
function getOperandShape(operand: BroadcastOperand): ArrayShape {
  return operand.type === "array" ? operand.shape : { rows: 1, cols: 1 };
}

/** The result shape: the largest operand size in each dimension. */
export function getBroadcastShape(shapes: ArrayShape[]): ArrayShape {
  return shapes.reduce<ArrayShape>(
    (largest, shape) => ({
      rows: Math.max(largest.rows, shape.rows),
      cols: Math.max(largest.cols, shape.cols),
    }),
    { rows: 1, cols: 1 }
  );
}

/**
 * Maps a cell of the broadcast result to the operand cell it reads. A dimension
 * of size 1 is repeated, so it always maps to index 0. Returns `undefined` when
 * the cell lies beyond the operand, which makes that result cell `#N/A`.
 */
export function mapBroadcastOffset(
  operandShape: ArrayShape,
  offset: Offset
): Offset | undefined {
  const x = operandShape.cols === 1 ? 0 : offset.x;
  const y = operandShape.rows === 1 ? 0 : offset.y;
  if (x >= operandShape.cols || y >= operandShape.rows) {
    return undefined;
  }
  return { x, y };
}

/** The area a broadcast result of `shape` occupies when spilled from `origin`. */
export function getBroadcastSpillArea(
  shape: ArrayShape,
  origin: CellAddress
): SpreadsheetRange {
  return {
    start: { col: origin.colIndex, row: origin.rowIndex },
    end: {
      col:
        shape.cols === Infinity
          ? { type: "infinity", sign: "positive" }
          : { type: "number", value: origin.colIndex + shape.cols - 1 },
      row:
        shape.rows === Infinity
          ? { type: "infinity", sign: "positive" }
          : { type: "number", value: origin.rowIndex + shape.rows - 1 },
    },
  };
}

function beyondOperandError(context: EvaluationContext): SingleEvaluationResult {
  return {
    type: "error",
    err: FormulaError.NA,
    message: "The array is smaller than the broadcast result",
    errAddress: context.dependencyNode,
  };
}

/**
 * Reads an operand's value for one cell of the broadcast result. Array operands
 * are read lazily through their `evaluate`, so the dependency lands on `readContext`.
 */
function readBroadcastOperandCell(
  operand: BroadcastOperand,
  offset: Offset,
  readContext: EvaluationContext,
  errorContext: EvaluationContext
): SingleEvaluationResult {
  if (operand.type === "scalar") {
    return operand.result;
  }
  const mapped = mapBroadcastOffset(operand.shape, offset);
  if (!mapped) {
    return beyondOperandError(errorContext);
  }
  return operand.result.evaluate(mapped, readContext);
}

/**
 * Creates the spilled result of an element-wise operation over `operands`, at
 * least one of which is an array.
 *
 * - The spill area starts at the origin and has the broadcast shape.
 * - `evaluate` computes one cell lazily with `combine`, reading operand cells
 *   through their own `evaluate`.
 * - `evaluateAllCells` is sparse, like the evaluateAllCells of a raw range:
 *   each array operand is read once with its own `evaluateAllCells` (a range is
 *   read through a range node and yields only populated cells), and a result
 *   cell is reported when some array operand has a cell there. Cells where every
 *   array operand is blank are omitted, as `A1:A3*2` omits the blank cells of
 *   `A1:A3`; single values never add cells. An array that would be repeated
 *   along an unbounded dimension (the `B1:D1` of `A:A*B1:D1`) adds no cells
 *   either, so an unbounded result stays as sparse as the operands that span
 *   the unbounded dimension. `#N/A` cells beyond a shorter operand are always
 *   reported when the requested block is bounded, and otherwise only where some
 *   operand has a cell. Within a reported cell a blank operand cell reads as
 *   the empty string, exactly as `evaluate` reads it.
 */
export function createBroadcastResult(options: {
  operands: BroadcastOperand[];
  combine: BroadcastCombine;
  /** The context the operands were evaluated in. */
  context: EvaluationContext;
  /** Debug description of the producer. */
  source: string;
}): SpilledValuesEvaluationResult {
  const { operands, combine, context, source } = options;
  const shape = getBroadcastShape(operands.map(getOperandShape));

  return {
    type: "spilled-values",
    source,
    spillArea: (origin) => getBroadcastSpillArea(shape, origin),
    evaluate: (offset, readContext) => {
      if (
        offset.x < 0 ||
        offset.y < 0 ||
        offset.x >= shape.cols ||
        offset.y >= shape.rows
      ) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: `${source}: the requested cell is outside the result`,
          errAddress: context.dependencyNode,
        };
      }
      return combine((index) =>
        readBroadcastOperandCell(operands[index]!, offset, readContext, context)
      );
    },
    evaluateAllCells: function (allCellsOptions) {
      return evaluateAllBroadcastCells.call(this, {
        operands,
        shape,
        combine,
        context,
        allCellsOptions,
      });
    },
  };
}

type EvaluateAllCellsOptions = Parameters<
  SpilledValuesEvaluationResult["evaluateAllCells"]
>[0];

/**
 * The block of the result a consumer asked for, relative to the result's start.
 * `intersection` is in sheet coordinates, with the result spilled from `origin`.
 */
function getRequestedRegion(
  shape: ArrayShape,
  origin: CellAddress,
  intersection: SpreadsheetRange | undefined
): Region | undefined {
  let x0 = 0;
  let y0 = 0;
  let x1 = shape.cols - 1;
  let y1 = shape.rows - 1;
  if (intersection) {
    x0 = Math.max(x0, intersection.start.col - origin.colIndex);
    y0 = Math.max(y0, intersection.start.row - origin.rowIndex);
    if (intersection.end.col.type === "number") {
      x1 = Math.min(x1, intersection.end.col.value - origin.colIndex);
    }
    if (intersection.end.row.type === "number") {
      y1 = Math.min(y1, intersection.end.row.value - origin.rowIndex);
    }
  }
  if (x0 > x1 || y0 > y1) {
    return undefined;
  }
  return { x0, x1, y0, y1 };
}

/** The operand cells that the result cells of `region` read. */
function getOperandRegion(
  operandShape: ArrayShape,
  region: Region
): Region | undefined {
  const x0 = operandShape.cols === 1 ? 0 : region.x0;
  const x1 =
    operandShape.cols === 1 ? 0 : Math.min(region.x1, operandShape.cols - 1);
  const y0 = operandShape.rows === 1 ? 0 : region.y0;
  const y1 =
    operandShape.rows === 1 ? 0 : Math.min(region.y1, operandShape.rows - 1);
  if (x0 > x1 || y0 > y1) {
    return undefined;
  }
  return { x0, x1, y0, y1 };
}

function isInRegion(offset: Offset, region: Region): boolean {
  return (
    offset.x >= region.x0 &&
    offset.x <= region.x1 &&
    offset.y >= region.y0 &&
    offset.y <= region.y1
  );
}

/**
 * Reads the populated cells of an array operand inside `operandRegion`, with
 * positions relative to the operand's start. Cells reported outside the block
 * are dropped.
 *
 * The operand's own `evaluate` is passed on: producers such as SEQUENCE call the
 * `evaluate` they receive, so forwarding the caller's would apply this operation
 * twice.
 *
 * A block that does not cover the whole operand is passed as the intersection,
 * so an unbounded operand such as `SEQUENCE(INFINITY)` is only read where it is
 * needed. Producers report positions relative to the start of their spill area
 * (the origin), also within an intersection.
 */
function readOperandCells(
  this: FormulaEvaluator,
  operand: Extract<BroadcastOperand, { type: "array" }>,
  operandRegion: Region,
  allCellsOptions: EvaluateAllCellsOptions
): EvaluateAllCellsResult {
  const { origin } = allCellsOptions;
  const coversOperand =
    operandRegion.x0 === 0 &&
    operandRegion.y0 === 0 &&
    operandRegion.x1 === operand.shape.cols - 1 &&
    operandRegion.y1 === operand.shape.rows - 1;

  const intersection: SpreadsheetRange | undefined = coversOperand
    ? undefined
    : {
        start: {
          col: origin.colIndex + operandRegion.x0,
          row: origin.rowIndex + operandRegion.y0,
        },
        end: {
          col:
            operandRegion.x1 === Infinity
              ? { type: "infinity", sign: "positive" }
              : { type: "number", value: origin.colIndex + operandRegion.x1 },
          row:
            operandRegion.y1 === Infinity
              ? { type: "infinity", sign: "positive" }
              : { type: "number", value: origin.rowIndex + operandRegion.y1 },
        },
      };

  const cells = operand.result.evaluateAllCells.call(this, {
    context: allCellsOptions.context,
    origin,
    lookupOrder: allCellsOptions.lookupOrder,
    evaluate: operand.result.evaluate,
    intersection,
  });
  if (cells.type !== "values") {
    return cells;
  }
  return {
    type: "values",
    values: cells.values.filter((cell) =>
      isInRegion(cell.relativePos, operandRegion)
    ),
  };
}

/**
 * Keys the cells of a block, relative to the array's start. A block of finite
 * width gets numeric keys, which are much faster than text keys.
 */
function createCellKey(
  region: Region
): (x: number, y: number) => number | string {
  const width = region.x1 - region.x0 + 1;
  if (width === Infinity) {
    return (x, y) => `${x},${y}`;
  }
  return (x, y) => (y - region.y0) * width + (x - region.x0);
}

/** The populated cells read from one array operand. */
type OperandCells = {
  list: CellInRangeResult[];
  key: (x: number, y: number) => number | string;
  byPosition: Map<number | string, SingleEvaluationResult>;
};

/**
 * A bounded block of at most this many cells tracks its reported cells in a
 * bitmap and is walked in lookup order. Larger and unbounded blocks track each
 * reported cell and sort them.
 */
const MAX_BITMAP_CELLS = 1 << 24;

function evaluateAllBroadcastCells(
  this: FormulaEvaluator,
  {
    operands,
    shape,
    combine,
    context,
    allCellsOptions,
  }: {
    operands: BroadcastOperand[];
    shape: ArrayShape;
    combine: BroadcastCombine;
    context: EvaluationContext;
    allCellsOptions: EvaluateAllCellsOptions;
  }
): EvaluateAllCellsResult {
  const region = getRequestedRegion(
    shape,
    allCellsOptions.origin,
    allCellsOptions.intersection
  );
  if (!region) {
    return { type: "values", values: [] };
  }

  // Fast path: one array spans the whole result and every other operand is a
  // single value (`A1:A3*2`). Its cells are the result cells, in its order.
  const arrayOperands = operands.filter((operand) => operand.type === "array");
  const [onlyArray] = arrayOperands;
  if (
    arrayOperands.length === 1 &&
    onlyArray &&
    onlyArray.shape.rows === shape.rows &&
    onlyArray.shape.cols === shape.cols
  ) {
    const cells = readOperandCells.call(this, onlyArray, region, allCellsOptions);
    if (cells.type !== "values") {
      return cells;
    }
    return {
      type: "values",
      values: cells.values.map(
        (cell): CellInRangeResult => ({
          relativePos: cell.relativePos,
          result: combine((index) => {
            const operand = operands[index]!;
            return operand.type === "scalar" ? operand.result : cell.result;
          }),
        })
      ),
    };
  }

  // Read every array operand once.
  const operandCells: (OperandCells | undefined)[] = [];
  for (let index = 0; index < operands.length; index++) {
    const operand = operands[index]!;
    if (operand.type === "scalar") {
      continue;
    }
    const operandRegion = getOperandRegion(operand.shape, region);
    if (!operandRegion) {
      // The requested block lies beyond this operand: all of it is #N/A.
      continue;
    }
    const cells = readOperandCells.call(
      this,
      operand,
      operandRegion,
      allCellsOptions
    );
    if (cells.type !== "values") {
      return cells;
    }
    const key = createCellKey(operandRegion);
    const byPosition = new Map<number | string, SingleEvaluationResult>();
    for (const cell of cells.values) {
      byPosition.set(key(cell.relativePos.x, cell.relativePos.y), cell.result);
    }
    operandCells[index] = { list: cells.values, key, byPosition };
  }

  // Computes the result cell at (x, y). The reader is shared by all cells, so
  // no closure is created per cell.
  let currentX = 0;
  let currentY = 0;
  const operandCell = (index: number): SingleEvaluationResult => {
    const operand = operands[index]!;
    if (operand.type === "scalar") {
      return operand.result;
    }
    const x = operand.shape.cols === 1 ? 0 : currentX;
    const y = operand.shape.rows === 1 ? 0 : currentY;
    if (x >= operand.shape.cols || y >= operand.shape.rows) {
      return beyondOperandError(context);
    }
    const read = operandCells[index];
    return read?.byPosition.get(read.key(x, y)) ?? BLANK;
  };
  const cellAt = (x: number, y: number): CellInRangeResult => {
    currentX = x;
    currentY = y;
    return { relativePos: { x, y }, result: combine(operandCell) };
  };

  const width = region.x1 - region.x0 + 1;
  const height = region.y1 - region.y0 + 1;
  const values =
    width * height <= MAX_BITMAP_CELLS
      ? collectBoundedCells({
          operands,
          operandCells,
          shape,
          region,
          lookupOrder: allCellsOptions.lookupOrder,
          cellAt,
        })
      : collectSparseCells({
          operands,
          operandCells,
          shape,
          region,
          lookupOrder: allCellsOptions.lookupOrder,
          cellAt,
        });
  return { type: "values", values };
}

type CollectCellsOptions = {
  operands: BroadcastOperand[];
  operandCells: (OperandCells | undefined)[];
  shape: ArrayShape;
  region: Region;
  lookupOrder: EvaluateAllCellsOptions["lookupOrder"];
  cellAt: (x: number, y: number) => CellInRangeResult;
};

/** Whether an operand of size 1 in a dimension is repeated along it. */
function getRepetition(operandShape: ArrayShape, shape: ArrayShape) {
  return {
    repeatsColumns: operandShape.cols === 1 && shape.cols > 1,
    repeatsRows: operandShape.rows === 1 && shape.rows > 1,
  };
}

/**
 * Reports the cells of a bounded block: marks them in a bitmap, then walks the
 * block in lookup order.
 */
function collectBoundedCells({
  operands,
  operandCells,
  shape,
  region,
  lookupOrder,
  cellAt,
}: CollectCellsOptions): CellInRangeResult[] {
  const width = region.x1 - region.x0 + 1;
  const height = region.y1 - region.y0 + 1;
  const marked = new Uint8Array(width * height);
  const fullRows = new Uint8Array(height);
  const fullColumns = new Uint8Array(width);
  let everything = false;

  operands.forEach((operand, index) => {
    if (operand.type === "scalar") {
      return;
    }
    const { repeatsColumns, repeatsRows } = getRepetition(operand.shape, shape);
    for (const { relativePos } of operandCells[index]?.list ?? []) {
      if (repeatsColumns && repeatsRows) {
        everything = true;
      } else if (repeatsColumns) {
        fullRows[relativePos.y - region.y0] = 1;
      } else if (repeatsRows) {
        fullColumns[relativePos.x - region.x0] = 1;
      } else {
        marked[
          (relativePos.y - region.y0) * width + (relativePos.x - region.x0)
        ] = 1;
      }
    }
    // The cells beyond a shorter operand are #N/A.
    if (operand.shape.cols > 1 && operand.shape.cols < shape.cols) {
      for (let x = Math.max(operand.shape.cols, region.x0); x <= region.x1; x++) {
        fullColumns[x - region.x0] = 1;
      }
    }
    if (operand.shape.rows > 1 && operand.shape.rows < shape.rows) {
      for (let y = Math.max(operand.shape.rows, region.y0); y <= region.y1; y++) {
        fullRows[y - region.y0] = 1;
      }
    }
  });

  const values: CellInRangeResult[] = [];
  const visit = (x: number, y: number) => {
    const column = x - region.x0;
    const row = y - region.y0;
    if (
      everything ||
      fullRows[row] === 1 ||
      fullColumns[column] === 1 ||
      marked[row * width + column] === 1
    ) {
      values.push(cellAt(x, y));
    }
  };
  if (lookupOrder === "row-major") {
    for (let y = region.y0; y <= region.y1; y++) {
      for (let x = region.x0; x <= region.x1; x++) {
        visit(x, y);
      }
    }
  } else {
    for (let x = region.x0; x <= region.x1; x++) {
      for (let y = region.y0; y <= region.y1; y++) {
        visit(x, y);
      }
    }
  }
  return values;
}

/**
 * Reports the cells of an unbounded (or very large) block: collects each
 * reported cell, then sorts them in lookup order.
 */
function collectSparseCells({
  operands,
  operandCells,
  shape,
  region,
  lookupOrder,
  cellAt,
}: CollectCellsOptions): CellInRangeResult[] {
  const positionKey = createCellKey(region);
  const positions = new Map<number | string, Offset>();
  // Adds every cell of a bounded block of the result.
  const addBlock = (x0: number, x1: number, y0: number, y1: number) => {
    for (let y = y0; y <= y1; y++) {
      for (let x = x0; x <= x1; x++) {
        const key = positionKey(x, y);
        if (!positions.has(key)) {
          positions.set(key, { x, y });
        }
      }
    }
  };
  const regionBounded = region.x1 !== Infinity && region.y1 !== Infinity;

  operands.forEach((operand, index) => {
    if (operand.type === "scalar") {
      return;
    }
    const { repeatsColumns, repeatsRows } = getRepetition(operand.shape, shape);
    // Repeating a cell along an unbounded dimension would report infinitely
    // many cells: such an operand adds no cells, like a single value.
    const addsCells =
      !(repeatsColumns && region.x1 === Infinity) &&
      !(repeatsRows && region.y1 === Infinity);
    if (addsCells) {
      for (const { relativePos } of operandCells[index]?.list ?? []) {
        addBlock(
          repeatsColumns ? region.x0 : relativePos.x,
          repeatsColumns ? region.x1 : relativePos.x,
          repeatsRows ? region.y0 : relativePos.y,
          repeatsRows ? region.y1 : relativePos.y
        );
      }
    }
    // The cells beyond a shorter operand are #N/A. Along an unbounded
    // dimension there are infinitely many: they are reported only where
    // another operand adds a cell.
    if (regionBounded) {
      if (operand.shape.cols > 1 && operand.shape.cols < shape.cols) {
        const start = Math.max(operand.shape.cols, region.x0);
        if (start <= region.x1) {
          addBlock(start, region.x1, region.y0, region.y1);
        }
      }
      if (operand.shape.rows > 1 && operand.shape.rows < shape.rows) {
        const start = Math.max(operand.shape.rows, region.y0);
        if (start <= region.y1) {
          addBlock(region.x0, region.x1, start, region.y1);
        }
      }
    }
  });

  const ordered = Array.from(positions.values());
  if (lookupOrder === "row-major") {
    ordered.sort((a, b) => a.y - b.y || a.x - b.x);
  } else {
    ordered.sort((a, b) => a.x - b.x || a.y - b.y);
  }
  return ordered.map((position) => cellAt(position.x, position.y));
}
