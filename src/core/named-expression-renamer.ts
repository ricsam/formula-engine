import type { NamedExpressionNode } from "../parser/ast";
import { astToString } from "../parser/formatter";
import { parseFormula } from "../parser/parser";
import {
  forEachNameWithLetScopes,
  getLetDeclaredName,
  mapNamesWithLetScopes,
  normalizeLetName,
} from "./ast-traverser";

/**
 * Renames named expression references in a formula string.
 *
 * Only references to the defined name change. LET declarations and the
 * references to them are local names, even when they are spelled like the
 * defined name, and stay as they are.
 *
 * A renamed reference must not be captured by a LET name: renaming the
 * defined name `y` to `x` in `LET(x, 1, x + y)` would otherwise turn the
 * reference into the local `x`. Every LET name in scope at such a reference
 * that matches the new name is therefore renamed too, together with its uses,
 * to `<name>_<n>`: the first such spelling that no name in the formula uses.
 * The underscore keeps it a name rather than a cell reference (`x1` is a cell,
 * `x_1` a name). The example becomes `LET(x_1, 1, x_1 + x)`.
 *
 * @param formula - The formula string (without the leading =)
 * @param oldName - The current named expression name to replace
 * @param newName - The new named expression name
 * @returns The updated formula string, or the original if it cannot be parsed
 */
export function renameNamedExpressionInFormula(
  formula: string,
  oldName: string,
  newName: string
): string {
  try {
    const ast = parseFormula(formula);

    // Find the LET names that would capture a renamed reference, including
    // the outer names a capturing inner name shadows.
    const newKey = normalizeLetName(newName);
    const capturing = new Set<NamedExpressionNode>();
    forEachNameWithLetScopes(ast, (node, role, scope) => {
      if (
        role.kind !== "defined" ||
        node.name !== oldName ||
        getLetDeclaredName(node) === undefined
      ) {
        return;
      }
      for (
        let binding = scope.get(newKey);
        binding;
        binding = binding.shadows
      ) {
        capturing.add(binding.declaration);
      }
    });

    const freshNames = new Map<NamedExpressionNode, string>();
    if (capturing.size > 0) {
      const usedNames = new Set([normalizeLetName(oldName), newKey]);
      forEachNameWithLetScopes(ast, (node) => {
        usedNames.add(normalizeLetName(node.name));
      });
      for (const declaration of capturing) {
        freshNames.set(declaration, freshLetName(declaration.name, usedNames));
      }
    }

    const transformedAST = mapNamesWithLetScopes(ast, (node, role) => {
      if (role.kind === "defined") {
        return node.name === oldName ? { ...node, name: newName } : node;
      }
      const freshName = freshNames.get(role.binding.declaration);
      return freshName === undefined ? node : { ...node, name: freshName };
    });

    return astToString(transformedAST);
  } catch (error) {
    // If parsing fails, return the original formula
    return formula;
  }
}

/** `<name>_<n>` for the smallest n whose spelling is not in `usedNames`. */
function freshLetName(name: string, usedNames: Set<string>): string {
  for (let suffix = 1; ; suffix++) {
    const candidate = `${name}_${suffix}`;
    const key = normalizeLetName(candidate);
    if (!usedNames.has(key)) {
      usedNames.add(key);
      return candidate;
    }
  }
}

/**
 * Checks if a formula contains references to a specific named expression.
 * LET names spelled like it are not references to it.
 * @param formula - The formula string (without the leading =)
 * @param namedExpressionName - The named expression name to search for
 * @returns True if the formula contains references to the named expression
 */
export function formulaReferencesNamedExpression(
  formula: string,
  namedExpressionName: string
): boolean {
  return getReferencedNamedExpressionNames(formula).includes(
    namedExpressionName
  );
}

/**
 * Gets all named expression names referenced in a formula. LET names and the
 * references to them are local to the formula and not included.
 * @param formula - The formula string (without the leading =)
 * @returns Array of unique named expression names referenced in the formula
 */
export function getReferencedNamedExpressionNames(formula: string): string[] {
  try {
    const ast = parseFormula(formula);
    const namedExpressionNames = new Set<string>();

    forEachNameWithLetScopes(ast, (node, role) => {
      if (role.kind === "defined" && node.name) {
        namedExpressionNames.add(node.name);
      }
    });

    return Array.from(namedExpressionNames);
  } catch (error) {
    // If parsing fails, return empty array
    return [];
  }
}
