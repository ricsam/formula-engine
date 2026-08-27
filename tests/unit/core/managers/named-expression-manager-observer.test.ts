import { describe, expect, test } from "bun:test";
import { serialize } from "../../../../src/core/map-serializer";
import {
  NamedExpressionManager,
  type NamedExpressionMutation,
} from "../../../../src/core/managers/named-expression-manager";
import type { NamedExpression } from "../../../../src/core/types";

function expression(name: string, value: string): NamedExpression {
  return { name, expression: value };
}

class CountingMap<TKey, TValue> extends Map<TKey, TValue> {
  clearCount = 0;

  override clear(): void {
    this.clearCount++;
    super.clear();
  }
}

describe("NamedExpressionManager mutation observer", () => {
  test("skips delta construction while observation is inactive", () => {
    let active = false;
    const batches: Array<readonly NamedExpressionMutation[]> = [];
    const manager = new NamedExpressionManager(
      (changes) => batches.push(changes),
      () => active
    );

    manager.addNamedExpression({ expressionName: "Ignored", expression: "1" });
    expect(batches).toHaveLength(0);

    active = true;
    manager.addNamedExpression({ expressionName: "Captured", expression: "2" });
    expect(batches.flat()).toHaveLength(1);
  });

  test("add, update, rename, and remove emit one detached entry delta", () => {
    const batches: Array<readonly NamedExpressionMutation[]> = [];
    const manager = new NamedExpressionManager((changes) => {
      batches.push(changes);
    });

    manager.addNamedExpression({ expressionName: "Rate", expression: "1" });
    expect(batches).toHaveLength(1);
    expect(batches[0]).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "named-expression",
      before: undefined,
      after: {
        expressionName: "Rate",
        expression: { name: "Rate", expression: "1" },
        index: 0,
      },
    });

    const addDelta = batches[0]?.[0];
    batches.length = 0;
    manager.updateNamedExpression({
      expressionName: "Rate",
      expression: "2",
    });
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "named-expression",
      before: { expression: { expression: "1" }, index: 0 },
      after: { expression: { expression: "2" }, index: 0 },
    });
    expect(addDelta).toMatchObject({
      after: { expression: { expression: "1" } },
    });

    batches.length = 0;
    manager.renameNamedExpression({
      expressionName: "Rate",
      newName: "TaxRate",
    });
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "named-expression",
      before: { expressionName: "Rate", index: 0 },
      after: {
        expressionName: "TaxRate",
        expression: { name: "TaxRate", expression: "2" },
        index: 0,
      },
    });

    batches.length = 0;
    expect(manager.removeNamedExpression({ expressionName: "TaxRate" })).toBe(
      true
    );
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "named-expression",
      before: { expressionName: "TaxRate", index: 0 },
      after: undefined,
    });
  });

  test("updateAllNamedExpressions emits one atomic bulk patch", () => {
    let capture = false;
    const batchSizes: number[] = [];
    const manager = new NamedExpressionManager((changes) => {
      if (capture) {
        batchSizes.push(changes.length);
      }
    });
    for (let index = 0; index < 1_030; index++) {
      manager.addNamedExpression({
        expressionName: `Name${index}`,
        expression: `${index}`,
      });
    }

    capture = true;
    const changed = manager.updateAllNamedExpressions(
      (formula) => `${formula}+1`
    );

    expect(changed).toHaveLength(1_030);
    expect(batchSizes).toEqual([1_030]);
  });

  test("bulk history replay rebuilds an affected expression map once", () => {
    const patches: NamedExpressionMutation[][] = [];
    let capture = false;
    const manager = new NamedExpressionManager((changes) => {
      if (capture) {
        patches.push([...changes]);
      }
    });
    const expressions = new CountingMap<string, NamedExpression>();
    manager.globalExpressions = expressions;

    for (let index = 0; index < 5_000; index++) {
      manager.addNamedExpression({
        expressionName: `Name${index}`,
        expression: `${index}`,
      });
    }
    const expectedOrder = Array.from(expressions.keys());

    capture = true;
    manager.updateAllNamedExpressions((formula) => `${formula}+1`);
    capture = false;
    expect(patches).toHaveLength(1);
    expect(patches[0]).toHaveLength(5_000);

    manager.applyHistoryChanges(patches[0]!, "undo");
    expect(expressions.clearCount).toBe(1);
    expect(Array.from(expressions.keys())).toEqual(expectedOrder);
    expect(expressions.get("Name4999")?.expression).toBe("4999");

    manager.applyHistoryChanges(patches[0]!, "redo");
    expect(expressions.clearCount).toBe(2);
    expect(Array.from(expressions.keys())).toEqual(expectedOrder);
    expect(expressions.get("Name4999")?.expression).toBe("4999+1");
  });

  test("reset deltas replay exact ordering and empty scope buckets", () => {
    let capture = false;
    const captured: NamedExpressionMutation[] = [];
    const manager = new NamedExpressionManager((changes) => {
      if (capture) {
        captured.push(...changes);
      }
    });
    manager.addWorkbook("Book1");
    manager.addWorkbook("Book2");
    manager.addSheet({ workbookName: "Book1", sheetName: "Data" });
    manager.addSheet({ workbookName: "Book1", sheetName: "Empty" });
    manager.addNamedExpression({ expressionName: "GlobalA", expression: "1" });
    manager.addNamedExpression({
      workbookName: "Book1",
      expressionName: "WorkbookA",
      expression: "2",
    });
    manager.addNamedExpression({
      workbookName: "Book1",
      sheetName: "Data",
      expressionName: "SheetA",
      expression: "3",
    });
    const before = serialize(manager.toSnapshot());

    const replacement = {
      globalExpressions: new Map([
        ["GlobalB", expression("GlobalB", "4")],
        ["GlobalA", expression("GlobalA", "10")],
      ]),
      workbookExpressions: new Map([
        ["Book2", new Map<string, NamedExpression>()],
        ["Book1", new Map([["WorkbookB", expression("WorkbookB", "5")]])],
        ["Book3", new Map<string, NamedExpression>()],
      ]),
      sheetExpressions: new Map([
        ["Book2", new Map<string, Map<string, NamedExpression>>()],
        [
          "Book1",
          new Map([
            ["Empty", new Map<string, NamedExpression>()],
            ["Data", new Map([["SheetB", expression("SheetB", "6")]])],
          ]),
        ],
        ["Book3", new Map<string, Map<string, NamedExpression>>()],
      ]),
    };

    capture = true;
    manager.resetNamedExpressions(replacement);
    capture = false;
    const after = serialize(manager.toSnapshot());

    expect(captured.length).toBeGreaterThan(0);
    expect(
      captured.every((change) =>
        change.kind === "named-expression"
          ? !(
              "globalExpressions" in (change.before?.expression ?? {}) ||
              "globalExpressions" in (change.after?.expression ?? {})
            )
          : true
      )
    ).toBe(true);

    manager.applyHistoryChanges(captured, "undo");
    expect(serialize(manager.toSnapshot())).toBe(before);

    manager.applyHistoryChanges(captured, "redo");
    expect(serialize(manager.toSnapshot())).toBe(after);
    expect(Array.from(manager.workbookExpressions.keys())).toEqual([
      "Book2",
      "Book1",
      "Book3",
    ]);
    expect(Array.from(manager.sheetExpressions.get("Book1")!.keys())).toEqual([
      "Empty",
      "Data",
    ]);
    expect(manager.workbookExpressions.get("Book3")?.size).toBe(0);
    expect(manager.sheetExpressions.get("Book3")?.size).toBe(0);
  });
});
