import { describe, expect, test } from "bun:test";
import { parseFormula } from "../parser/parser";
import { astToString } from "../parser/formatter";
import {
  forEachNameWithLetScopes,
  mapNamesWithLetScopes,
  type NameRole,
} from "./ast-traverser";

// traverseAST, findNodesByType and transformAST are covered in
// tests/unit/core/ast-traverser.test.ts; these cases cover LET scoping.

/** Each name in the formula as `name:role`, where a local names its declaration's argument index. */
function roles(formula: string): string[] {
  const result: string[] = [];
  forEachNameWithLetScopes(parseFormula(formula), (node, role: NameRole) => {
    const text = astToString(node);
    if (role.kind === "defined") result.push(`${text}:defined`);
    else result.push(`${text}:${role.kind}@${role.binding.argIndex}`);
  });
  return result;
}

describe("LET scoping", () => {
  test("classifies declarations, local references and defined names", () => {
    expect(roles("LET(x, y, z, x + 1, x * z + w)")).toEqual([
      "x:declaration@0",
      "y:defined",
      "z:declaration@2",
      "x:local@0",
      "x:local@0",
      "z:local@2",
      "w:defined",
    ]);
  });

  test("a value does not see its own name", () => {
    expect(roles("LET(x, x, x)")).toEqual([
      "x:declaration@0",
      "x:defined",
      "x:local@0",
    ]);
  });

  test("names are case-insensitive and qualified names are never local", () => {
    expect(roles("LET(Total, 1, TOTAL + Sheet1!Total)")).toEqual([
      "Total:declaration@0",
      "TOTAL:local@0",
      "Sheet1!Total:defined",
    ]);
  });

  test("nested LETs shadow outer names and remember what they shadow", () => {
    const shadowed: Array<number | undefined> = [];
    forEachNameWithLetScopes(
      parseFormula("LET(x, 1, LET(x, 2, x) + x)"),
      (node, role) => {
        if (role.kind === "local") shadowed.push(role.binding.shadows?.argIndex);
      }
    );
    expect(roles("LET(x, 1, LET(x, 2, x) + x)")).toEqual([
      "x:declaration@0",
      "x:declaration@0",
      "x:local@0",
      "x:local@0",
    ]);
    // The inner local shadows the outer declaration; the outer local shadows nothing.
    expect(shadowed).toEqual([0, undefined]);
  });

  test("walks malformed LET calls leniently", () => {
    // A cell reference in a name position declares nothing.
    expect(roles("LET(x1, y, x1)")).toEqual(["y:defined"]);
    // Without a calculation, every pair still declares its name.
    expect(roles("LET(a, 1, b, a)")).toEqual([
      "a:declaration@0",
      "b:declaration@2",
      "a:local@0",
    ]);
    expect(roles("LET(a)")).toEqual(["a:defined"]);
  });

  test("names outside a LET are defined names", () => {
    expect(roles("LET(x, 1, x) + x")).toEqual([
      "x:declaration@0",
      "x:local@0",
      "x:defined",
    ]);
  });

  test("mapNamesWithLetScopes rebuilds the tree with mapped names", () => {
    const mapped = mapNamesWithLetScopes(
      parseFormula("SUM(LET(x, a, x), {a, b})"),
      (node, role) =>
        role.kind === "defined" ? { ...node, name: node.name.toUpperCase() } : node
    );
    expect(astToString(mapped)).toBe("SUM(LET(x,A,x),{A,B})");
  });
});
