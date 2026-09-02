import { describe, expect, test } from "bun:test";

import {
  functionDefinitions,
  functions,
} from "../functions/function-registry";
import {
  findFormulaFunction,
  getFormulaFunctionCatalog,
} from "./function-catalog";

describe("formula function catalog", () => {
  test("describes every canonical registered function in name order", () => {
    const catalog = getFormulaFunctionCatalog();
    const registeredNames = Object.keys(functionDefinitions).sort((left, right) =>
      left.localeCompare(right)
    );

    expect(catalog.map(({ name }) => name)).toEqual(registeredNames);
    expect(catalog).toHaveLength(38);

    for (const descriptor of catalog) {
      expect(descriptor.description.length).toBeGreaterThan(0);
      expect(descriptor.signature.startsWith(`${descriptor.name}(`)).toBe(true);
      expect(descriptor.parameters.every(({ name }) => name.length > 0)).toBe(
        true
      );
    }
  });

  test("finds canonical names and aliases case-insensitively", () => {
    const canonical = findFormulaFunction("concatenate");

    expect(canonical).toBeDefined();
    expect(findFormulaFunction("CoNcAt")).toBe(canonical!);
    expect(canonical?.aliases).toEqual(["CONCAT"]);
    expect(findFormulaFunction("missing")).toBeUndefined();
  });

  test("exposes structured arguments for completion snippets", () => {
    expect(findFormulaFunction("SUMIFS")).toMatchObject({
      category: "math",
      signature:
        "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
      parameters: [
        { name: "sum_range", optional: false, repeatable: false },
        { name: "criteria_range1", optional: false, repeatable: false },
        { name: "criteria1", optional: false, repeatable: false },
        { name: "criteria_range2", optional: true, repeatable: true },
        { name: "criteria2", optional: true, repeatable: true },
      ],
    });
  });

  test("resolves every callable registry name to canonical metadata", () => {
    for (const [callableName, definition] of Object.entries(functions)) {
      expect(findFormulaFunction(callableName)?.name).toBe(definition.name);
    }
  });

  test("returns deeply immutable catalog data", () => {
    const catalog = getFormulaFunctionCatalog();
    const first = catalog[0]!;

    expect(Object.isFrozen(catalog)).toBe(true);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first.aliases)).toBe(true);
    expect(Object.isFrozen(first.parameters)).toBe(true);
    expect(first.parameters.every(Object.isFrozen)).toBe(true);
  });
});
