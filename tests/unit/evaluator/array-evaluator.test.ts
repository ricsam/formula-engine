import { test, expect, describe } from "bun:test";
import {
  getArrayDimensions,
  to2DArray,
  flatten,
  areDimensionsCompatible,
  broadcast,
  broadcastToSize,
  elementWiseBinaryOp,
  elementWiseUnaryOp,
  reduceArray,
  constrainArray,
  transpose,
  coerceToBoolean,
  willSpill,
  calculateSpillRange,
  arraySum,
  arrayProduct,
  arrayCount,
  ArrayOperationCache,
} from "../../../src/evaluator/array-evaluator";
import { FormulaError, type CellValue } from "../../../src/core/types";

describe("Array Evaluator", () => {
  describe("getArrayDimensions", () => {
    test("should get dimensions of scalar", () => {
      expect(getArrayDimensions(42)).toEqual({ rows: 1, cols: 1 });
      expect(getArrayDimensions("text")).toEqual({ rows: 1, cols: 1 });
      expect(getArrayDimensions(true)).toEqual({ rows: 1, cols: 1 });
    });

    test("should get dimensions of 1D array", () => {
      expect(getArrayDimensions([1, 2, 3])).toEqual({ rows: 3, cols: 1 });
      expect(getArrayDimensions([])).toEqual({ rows: 0, cols: 1 });
    });

    test("should get dimensions of 2D array", () => {
      expect(
        getArrayDimensions([
          [1, 2],
          [3, 4],
          [5, 6],
        ])
      ).toEqual({ rows: 3, cols: 2 });
      expect(getArrayDimensions([[1]])).toEqual({ rows: 1, cols: 1 });
      expect(getArrayDimensions([[]])).toEqual({ rows: 1, cols: 0 });
    });
  });

  describe("to2DArray", () => {
    test("should convert scalar to 2D array", () => {
      expect(to2DArray(42)).toEqual([[42]]);
      expect(to2DArray("text")).toEqual([["text"]]);
      expect(to2DArray(undefined)).toEqual([[undefined]]);
    });

    test("should convert 1D array to 2D column vector", () => {
      expect(to2DArray([1, 2, 3])).toEqual([[1], [2], [3]]);
      expect(to2DArray([])).toEqual([]);
    });

    test("should keep 2D array unchanged", () => {
      const array = [
        [1, 2],
        [3, 4],
      ];
      expect(to2DArray(array)).toBe(array);
    });
  });

  describe("flatten", () => {
    test("should flatten 2D array", () => {
      expect(
        flatten([
          [1, 2],
          [3, 4],
          [5, 6],
        ])
      ).toEqual([1, 2, 3, 4, 5, 6]);
      expect(flatten([[1]])).toEqual([1]);
      expect(flatten([])).toEqual([]);
      expect(flatten([[]])).toEqual([]);
    });
  });

  describe("areDimensionsCompatible", () => {
    test("scalars are compatible with everything", () => {
      const scalar = { rows: 1, cols: 1 };
      const array = { rows: 3, cols: 4 };

      expect(areDimensionsCompatible(scalar, array)).toBe(true);
      expect(areDimensionsCompatible(array, scalar)).toBe(true);
    });

    test("same dimensions are compatible", () => {
      const dim1 = { rows: 3, cols: 4 };
      const dim2 = { rows: 3, cols: 4 };

      expect(areDimensionsCompatible(dim1, dim2)).toBe(true);
    });

    test("broadcastable dimensions are compatible", () => {
      expect(
        areDimensionsCompatible({ rows: 3, cols: 1 }, { rows: 3, cols: 4 })
      ).toBe(true);

      expect(
        areDimensionsCompatible({ rows: 1, cols: 4 }, { rows: 3, cols: 4 })
      ).toBe(true);
    });

    test("incompatible dimensions", () => {
      expect(
        areDimensionsCompatible({ rows: 3, cols: 4 }, { rows: 2, cols: 4 })
      ).toBe(false);

      expect(
        areDimensionsCompatible({ rows: 3, cols: 4 }, { rows: 3, cols: 5 })
      ).toBe(false);
    });
  });

  describe("broadcast", () => {
    test("should broadcast scalar to array", () => {
      const result = broadcast(
        [[42]],
        [
          [1, 2],
          [3, 4],
        ]
      );

      expect(result).not.toBe("#VALUE!");
      if (typeof result !== "string") {
        expect(result.array1).toEqual([
          [42, 42],
          [42, 42],
        ]);
        expect(result.array2).toEqual([
          [1, 2],
          [3, 4],
        ]);
        expect(result.dimensions).toEqual({ rows: 2, cols: 2 });
      }
    });

    test("should broadcast column to matrix", () => {
      const result = broadcast(
        [[1], [2]],
        [
          [10, 20, 30],
          [40, 50, 60],
        ]
      );

      expect(result).not.toBe("#VALUE!");
      if (typeof result !== "string") {
        expect(result.array1).toEqual([
          [1, 1, 1],
          [2, 2, 2],
        ]);
        expect(result.dimensions).toEqual({ rows: 2, cols: 3 });
      }
    });

    test("should broadcast row to matrix", () => {
      const result = broadcast([[1, 2, 3]], [[10], [20], [30]]);

      expect(result).not.toBe("#VALUE!");
      if (typeof result !== "string") {
        expect(result.array1).toEqual([
          [1, 2, 3],
          [1, 2, 3],
          [1, 2, 3],
        ]);
        expect(result.dimensions).toEqual({ rows: 3, cols: 3 });
      }
    });

    test("should return error for incompatible arrays", () => {
      const result = broadcast(
        [[1, 2]],
        [
          [1, 2, 3],
          [4, 5, 6],
        ]
      );
      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe("broadcastToSize", () => {
    test("should broadcast scalar", () => {
      const result = broadcastToSize([[42]], { rows: 2, cols: 3 });
      expect(result).toEqual([
        [42, 42, 42],
        [42, 42, 42],
      ]);
    });

    test("should broadcast column", () => {
      const result = broadcastToSize([[1], [2]], { rows: 2, cols: 3 });
      expect(result).toEqual([
        [1, 1, 1],
        [2, 2, 2],
      ]);
    });

    test("should broadcast row", () => {
      const result = broadcastToSize([[1, 2, 3]], { rows: 3, cols: 3 });
      expect(result).toEqual([
        [1, 2, 3],
        [1, 2, 3],
        [1, 2, 3],
      ]);
    });

    test("should return error for invalid broadcast", () => {
      const result = broadcastToSize(
        [
          [1, 2],
          [3, 4],
        ],
        { rows: 3, cols: 3 }
      );
      expect(result).toBe(FormulaError.VALUE);
    });

    test("should handle empty arrays", () => {
      const result = broadcastToSize([[]], { rows: 2, cols: 2 });
      expect(result).toEqual([
        [undefined, undefined],
        [undefined, undefined],
      ]);
    });
  });

  describe("elementWiseBinaryOp", () => {
    test("should perform element-wise addition", () => {
      const add = (a: CellValue, b: CellValue) => (a as number) + (b as number);
      const result = elementWiseBinaryOp(
        [
          [1, 2],
          [3, 4],
        ],
        [
          [10, 20],
          [30, 40],
        ],
        add
      );

      expect(result).toEqual([
        [11, 22],
        [33, 44],
      ]);
    });

    test("should broadcast and operate", () => {
      const multiply = (a: CellValue, b: CellValue) =>
        (a as number) * (b as number);
      const result = elementWiseBinaryOp(
        [[2]],
        [
          [1, 2, 3],
          [4, 5, 6],
        ],
        multiply
      );

      expect(result).toEqual([
        [2, 4, 6],
        [8, 10, 12],
      ]);
    });

    test("should propagate errors", () => {
      const add = (a: CellValue, b: CellValue) => (a as number) + (b as number);
      const result = elementWiseBinaryOp(
        [
          [1, "#DIV/0!"],
          [3, 4],
        ],
        [
          [10, 20],
          [30, 40],
        ],
        add
      );

      expect(result).toEqual([
        [11, "#DIV/0!"],
        [33, 44],
      ]);
    });

    test("should return error for incompatible arrays", () => {
      const add = (a: CellValue, b: CellValue) => (a as number) + (b as number);
      const result = elementWiseBinaryOp(
        [
          [1, 2],
          [3, 4],
        ], // 2x2 array
        [
          [1, 2, 3],
          [4, 5, 6],
        ], // 2x3 array - incompatible column dimensions
        add
      );

      expect(result).toBe(FormulaError.VALUE);
    });
  });

  describe("elementWiseUnaryOp", () => {
    test("should perform element-wise negation", () => {
      const negate = (value: CellValue) => -(value as number);
      const result = elementWiseUnaryOp(
        [
          [1, -2],
          [3, -4],
        ],
        negate
      );

      expect(result).toEqual([
        [-1, 2],
        [-3, 4],
      ]);
    });

    test("should propagate errors", () => {
      const double = (value: CellValue) => (value as number) * 2;
      const result = elementWiseUnaryOp(
        [
          [1, "#REF!"],
          [3, 4],
        ],
        double
      );

      expect(result).toEqual([
        [2, "#REF!"],
        [6, 8],
      ]);
    });
  });

  describe("reduceArray", () => {
    test("should reduce to scalar", () => {
      const sum = (acc: CellValue, val: CellValue) =>
        (acc as number) + (val as number);
      const result = reduceArray(
        [
          [1, 2, 3],
          [4, 5, 6],
        ],
        sum,
        0,
        null
      );

      expect(result).toBe(21);
    });

    test("should reduce along rows (axis=0)", () => {
      const sum = (acc: CellValue, val: CellValue) =>
        (acc as number) + (val as number);
      const result = reduceArray(
        [
          [1, 2, 3],
          [4, 5, 6],
          [7, 8, 9],
        ],
        sum,
        0,
        0
      );

      expect(result).toEqual([[12, 15, 18]]);
    });

    test("should reduce along columns (axis=1)", () => {
      const sum = (acc: CellValue, val: CellValue) =>
        (acc as number) + (val as number);
      const result = reduceArray(
        [
          [1, 2, 3],
          [4, 5, 6],
          [7, 8, 9],
        ],
        sum,
        0,
        1
      );

      expect(result).toEqual([[6], [15], [24]]);
    });

    test("should propagate errors", () => {
      const sum = (acc: CellValue, val: CellValue) =>
        (acc as number) + (val as number);
      const result = reduceArray(
        [
          [1, "#VALUE!", 3],
          [4, 5, 6],
        ],
        sum,
        0,
        null
      );

      expect(result).toBe(FormulaError.VALUE);
    });

    test("should skip errors when reducing", () => {
      const sum = (acc: CellValue, val: CellValue) =>
        (acc as number) + (val as number);
      const result = reduceArray(
        [
          [1, 2, 3],
          [4, 5, 6],
        ],
        sum,
        0,
        null
      );

      expect(result).toBe(21);
    });
  });

  describe("constrainArray", () => {
    test("should constrain array size", () => {
      const result = constrainArray(
        [
          [1, 2, 3, 4],
          [5, 6, 7, 8],
          [9, 10, 11, 12],
        ],
        2,
        3
      );

      expect(result).toEqual([
        [1, 2, 3],
        [5, 6, 7],
      ]);
    });

    test("should handle smaller arrays", () => {
      const result = constrainArray(
        [
          [1, 2],
          [3, 4],
        ],
        5,
        5
      );

      expect(result).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    test("should handle empty arrays", () => {
      const result = constrainArray([], 2, 2);
      expect(result).toEqual([]);
    });
  });

  describe("transpose", () => {
    test("should transpose square matrix", () => {
      const result = transpose([
        [1, 2],
        [3, 4],
      ]);
      expect(result).toEqual([
        [1, 3],
        [2, 4],
      ]);
    });

    test("should transpose rectangular matrix", () => {
      const result = transpose([
        [1, 2, 3],
        [4, 5, 6],
      ]);
      expect(result).toEqual([
        [1, 4],
        [2, 5],
        [3, 6],
      ]);
    });

    test("should handle single row", () => {
      const result = transpose([[1, 2, 3]]);
      expect(result).toEqual([[1], [2], [3]]);
    });

    test("should handle single column", () => {
      const result = transpose([[1], [2], [3]]);
      expect(result).toEqual([[1, 2, 3]]);
    });

    test("should handle empty array", () => {
      const result = transpose([]);
      expect(result).toEqual([]);
    });

    test("should handle undefined values", () => {
      const result = transpose([
        [1, undefined],
        [undefined, 2],
      ]);
      expect(result).toEqual([
        [1, undefined],
        [undefined, 2],
      ]);
    });
  });

  describe("coerceToBoolean", () => {
    test("should handle boolean values", () => {
      expect(coerceToBoolean(true)).toBe(true);
      expect(coerceToBoolean(false)).toBe(false);
    });

    test("should handle numbers", () => {
      expect(coerceToBoolean(1)).toBe(true);
      expect(coerceToBoolean(-1)).toBe(true);
      expect(coerceToBoolean(0)).toBe(false);
    });

    test("should handle strings", () => {
      expect(coerceToBoolean("text")).toBe(true);
      expect(coerceToBoolean("")).toBe(false);
    });

    test("should handle undefined and errors", () => {
      expect(coerceToBoolean(undefined)).toBe(false);
      expect(coerceToBoolean(FormulaError.VALUE)).toBe(false);
    });
  });

  describe("willSpill", () => {
    test("should not spill single values", () => {
      expect(willSpill(5, 5, { rows: 1, cols: 1 }, 100, 100)).toBe(false);
    });

    test("should detect row spill", () => {
      expect(willSpill(99, 5, { rows: 3, cols: 1 }, 100, 100)).toBe(true);
    });

    test("should detect column spill", () => {
      expect(willSpill(5, 99, { rows: 1, cols: 3 }, 100, 100)).toBe(true);
    });

    test("should not spill when within bounds", () => {
      expect(willSpill(5, 5, { rows: 10, cols: 10 }, 100, 100)).toBe(false);
    });
  });

  describe("calculateSpillRange", () => {
    test("should calculate spill range", () => {
      const result = calculateSpillRange(5, 10, [
        [1, 2, 3],
        [4, 5, 6],
      ]);

      expect(result).toEqual({
        startRow: 5,
        startCol: 10,
        endRow: 6,
        endCol: 12,
      });
    });

    test("should handle single cell", () => {
      const result = calculateSpillRange(0, 0, [[42]]);

      expect(result).toEqual({
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      });
    });
  });

  describe("arraySum", () => {
    test("should sum numeric values", () => {
      expect(
        arraySum([
          [1, 2, 3],
          [4, 5, 6],
        ])
      ).toBe(21);
    });

    test("should treat booleans as numbers", () => {
      expect(arraySum([[true, false, true]])).toBe(2);
    });

    test("should skip non-numeric values", () => {
      expect(
        arraySum([
          [1, "text", 2],
          [undefined, 3, "#VALUE!"],
        ])
      ).toBe(6);
    });

    test("should return 0 for no numeric values", () => {
      expect(
        arraySum([
          ["a", "b"],
          ["c", "d"],
        ])
      ).toBe(0);
    });
  });

  describe("arrayProduct", () => {
    test("should multiply numeric values", () => {
      expect(arrayProduct([[2, 3, 4]])).toBe(24);
    });

    test("should treat booleans as numbers", () => {
      expect(arrayProduct([[2, true, 3, false]])).toBe(0);
    });

    test("should skip non-numeric values", () => {
      expect(
        arrayProduct([
          [2, "text", 3],
          [undefined, 4],
        ])
      ).toBe(24);
    });

    test("should return 1 for no numeric values", () => {
      expect(arrayProduct([["a", "b"]])).toBe(1);
    });
  });

  describe("arrayCount", () => {
    test("should count numeric values", () => {
      expect(
        arrayCount([
          [1, 2, 3],
          [4, 5, 6],
        ])
      ).toBe(6);
    });

    test("should count booleans", () => {
      expect(arrayCount([[true, false, true]])).toBe(3);
    });

    test("should skip non-numeric values", () => {
      expect(
        arrayCount([
          [1, "text", 2],
          [undefined, 3, "#VALUE!"],
        ])
      ).toBe(3);
    });

    test("should return 0 for no numeric values", () => {
      expect(
        arrayCount([
          ["a", "b"],
          [undefined, undefined],
        ])
      ).toBe(0);
    });
  });

  describe("ArrayOperationCache", () => {
    test("should cache and retrieve results", () => {
      const cache = new ArrayOperationCache();
      const result = [
        [1, 2],
        [3, 4],
      ];

      cache.set("sum", result, "A1:B2");
      expect(cache.get("sum", "A1:B2")).toEqual(result);
    });

    test("should return undefined for cache miss", () => {
      const cache = new ArrayOperationCache();
      expect(cache.get("sum", "A1:B2")).toBeUndefined();
    });

    test("should respect max size", () => {
      const cache = new ArrayOperationCache(2);

      cache.set("op1", [[1]], "arg1");
      cache.set("op2", [[2]], "arg2");
      cache.set("op3", [[3]], "arg3"); // Should evict op1

      expect(cache.get("op1", "arg1")).toBeUndefined();
      expect(cache.get("op2", "arg2")).toEqual([[2]]);
      expect(cache.get("op3", "arg3")).toEqual([[3]]);
    });

    test("should clear cache", () => {
      const cache = new ArrayOperationCache();

      cache.set("sum", [[42]], "A1");
      cache.clear();

      expect(cache.get("sum", "A1")).toBeUndefined();
    });
  });
});
