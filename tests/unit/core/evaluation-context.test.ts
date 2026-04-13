import { describe, expect, test } from "bun:test";
import {
  eligibleKeysForContext,
  getContextDependencyKey,
} from "../../../src/evaluator/evaluation-context";

describe("eligibleKeysForContext", () => {
  test("returns exact matches before broader wildcard matches", () => {
    const context = {
      workbookName: "Workbook",
      sheetName: "Sheet",
      tableName: "Table",
      rowIndex: 34,
      colIndex: 25,
    };

    const keys = eligibleKeysForContext(context);

    expect(keys[0]).toBe(getContextDependencyKey(context));
    expect(keys.at(-1)).toBe(
      getContextDependencyKey({
        workbookName: undefined,
        sheetName: undefined,
        tableName: undefined,
        rowIndex: undefined,
        colIndex: undefined,
      })
    );
  });
});
