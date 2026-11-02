import {
  getLetDeclaredName,
  isLetCall,
  normalizeLetName,
} from "../../../core/ast-traverser";
import { FormulaError } from "../../../core/types";
import type { ASTNode, FunctionNode, ValueNode } from "../../../parser/ast";
import { astToString } from "../../../parser/formatter";

/**
 * The largest number of AST nodes a LET may expand to.
 *
 * Expansion replaces every use of a name with the expression bound to it, so a
 * name used n times copies its expression n times. A chain of names that each
 * use the previous name twice (`b, a+a, c, b+b, ...`) doubles the expansion with
 * every pair: 20 such pairs expand to about a million nodes. The expanded
 * expression is evaluated like a written formula of that size: every
 * sub-expression is cached under its formula text, so evaluation time and
 * cache memory grow with the expanded size (and, through the cache keys, with
 * its depth), not with the length of the written formula. A LET past this
 * limit returns `#CALC!` instead of stalling the engine.
 *
 * 10,000 nodes is several times the largest formula Excel accepts (8,192
 * characters). At the limit, a bushy expansion such as the doubling chain
 * evaluates in well under 100 ms; a deliberately deep one (a long `+1+1+...`
 * chain over a large bound expression) costs about as much as the longest
 * plain formulas Excel accepts, around a second.
 */
export const MAX_LET_EXPANSION_NODES = 10_000;

/** A LET call that cannot be expanded; evaluation reports `formulaError`. */
export class LetExpansionError extends Error {
  constructor(
    readonly formulaError: FormulaError,
    message: string
  ) {
    super(message);
    this.name = "LetExpansionError";
  }
}

/**
 * An omitted name value (`LET(x, , x + 1)`) binds 0, like an omitted number
 * argument. This matches LibreOffice, whose LET tests are written against
 * Excel ("Unspecified value is defaulting to 0").
 */
const OMITTED_NAME_VALUE: ValueNode = {
  type: "value",
  value: { type: "number", value: 0 },
};

/** Expanded expressions bound to the LET names in scope, by normalized name. */
type Bindings = ReadonlyMap<string, ASTNode>;

/** Memoized node counts of (possibly shared) expanded subtrees. */
type Sizes = WeakMap<ASTNode, number>;

/**
 * Expand a LET call into an equivalent expression without LET names.
 *
 * Every unqualified name bound by this LET or a nested one is replaced with
 * the expression bound to it, so the result is ordinary, self-contained
 * formula AST: its text is a valid formula and fully determines its value.
 * That keeps the evaluator's AST cache (keyed by formula text and the context
 * the result depends on), dependency tracking, spills, `ROW()`-style context
 * dependencies and snapshots correct without any knowledge of LET.
 *
 * The substitution is hygienic. Bindings are processed in order, and the
 * expression bound to a name is inserted as it is, without being scanned
 * again, so a later or nested LET name can never capture a name inside it:
 * with a defined name `y`, `LET(x, y, LET(y, 2, x + y))` expands to `y + 2`
 * where `y` still means the defined name.
 *
 * Bound expressions are shared rather than copied, so expansion takes time and
 * memory proportional to the written formula. The expanded size is bounded by
 * {@link MAX_LET_EXPANSION_NODES}.
 *
 * @throws LetExpansionError when the call is malformed (`#VALUE!` for a wrong
 * argument count or a missing calculation, `#NAME?` for an invalid or repeated
 * name) or expands past the size limit (`#CALC!`). Nested LET calls are
 * validated as part of their enclosing LET.
 */
export function expandLet(node: FunctionNode): ASTNode {
  const sizes: Sizes = new WeakMap();
  const expanded = expandLetCall(node, new Map(), sizes);
  const size = sizeOf(expanded, sizes);
  if (size > MAX_LET_EXPANSION_NODES) {
    throw new LetExpansionError(
      FormulaError.CALC,
      `LET expands to ${formatCount(size)} nodes, more than the limit of ` +
        `${formatCount(MAX_LET_EXPANSION_NODES)}. Every use of a name ` +
        `repeats the expression bound to it; use names built from other ` +
        `names fewer times.`
    );
  }
  return expanded;
}

function expandLetCall(
  node: FunctionNode,
  outer: Bindings,
  sizes: Sizes
): ASTNode {
  const args = node.args;
  if (args.length < 3 || args.length % 2 === 0) {
    throw new LetExpansionError(
      FormulaError.VALUE,
      `LET expects one or more name/value pairs followed by a calculation ` +
        `(an odd number of at least 3 arguments), but got ${args.length}`
    );
  }

  const bindings = new Map(outer);
  const declaredHere = new Set<string>();
  for (let index = 0; index < args.length - 1; index += 2) {
    const nameArgument = args[index]!;
    const name = getLetDeclaredName(nameArgument);
    if (name === undefined) {
      throw new LetExpansionError(
        FormulaError.NAME,
        describeInvalidName(nameArgument, index / 2 + 1)
      );
    }

    const key = normalizeLetName(name);
    if (declaredHere.has(key)) {
      throw new LetExpansionError(
        FormulaError.NAME,
        `LET declares the name ${name} more than once`
      );
    }
    declaredHere.add(key);

    // The value sees the names declared before it (including an outer
    // binding of the same name), never its own.
    const valueArgument = args[index + 1]!;
    bindings.set(
      key,
      valueArgument.type === "empty"
        ? OMITTED_NAME_VALUE
        : substitute(valueArgument, bindings, sizes)
    );
  }

  const calculation = args[args.length - 1]!;
  if (calculation.type === "empty") {
    throw new LetExpansionError(
      FormulaError.VALUE,
      "LET is missing its final calculation"
    );
  }
  return substitute(calculation, bindings, sizes);
}

/**
 * Replace the free uses of bound names in `node`. Returns `node` itself when
 * nothing changes, so untouched subtrees stay shared.
 */
function substitute(node: ASTNode, bindings: Bindings, sizes: Sizes): ASTNode {
  switch (node.type) {
    case "named-expression": {
      const name = getLetDeclaredName(node);
      if (name === undefined) return node;
      return bindings.get(normalizeLetName(name)) ?? node;
    }

    case "function": {
      if (isLetCall(node)) return expandLetCall(node, bindings, sizes);
      const args = substituteEach(node.args, bindings, sizes);
      return args ? { ...node, args } : node;
    }

    case "binary-op": {
      const left = substitute(node.left, bindings, sizes);
      const right = substitute(node.right, bindings, sizes);
      return left === node.left && right === node.right
        ? node
        : { ...node, left, right };
    }

    case "unary-op": {
      const operand = substitute(node.operand, bindings, sizes);
      return operand === node.operand ? node : { ...node, operand };
    }

    case "array": {
      let changed = false;
      const elements = node.elements.map((row) => {
        const substituted = substituteEach(row, bindings, sizes);
        if (!substituted) return row;
        changed = true;
        return substituted;
      });
      return changed ? { ...node, elements } : node;
    }

    default:
      // References, ranges, literals and structured references hold no names.
      return node;
  }
}

function substituteEach(
  nodes: readonly ASTNode[],
  bindings: Bindings,
  sizes: Sizes
): ASTNode[] | undefined {
  let changed = false;
  const substituted = nodes.map((node) => {
    const result = substitute(node, bindings, sizes);
    if (result !== node) changed = true;
    return result;
  });
  return changed ? substituted : undefined;
}

/** The number of nodes in `node` as a tree, counting shared subtrees per use. */
function sizeOf(node: ASTNode, sizes: Sizes): number {
  const known = sizes.get(node);
  if (known !== undefined) return known;

  let size = 1;
  switch (node.type) {
    case "function":
      for (const arg of node.args) size += sizeOf(arg, sizes);
      break;
    case "binary-op":
      size += sizeOf(node.left, sizes) + sizeOf(node.right, sizes);
      break;
    case "unary-op":
      size += sizeOf(node.operand, sizes);
      break;
    case "array":
      for (const row of node.elements) {
        for (const element of row) size += sizeOf(element, sizes);
      }
      break;
    case "3d-range":
      size += sizeOf(node.reference, sizes);
      break;
  }
  sizes.set(node, size);
  return size;
}

function describeInvalidName(node: ASTNode, position: number): string {
  if (node.type === "empty") return `LET name ${position} is missing`;
  const text = astToString(node);
  switch (node.type) {
    case "reference":
    case "range":
      return `LET name ${position} (${text}) is a cell reference, not a name`;
    case "named-expression":
      return `LET name ${position} (${text}) must not be qualified by a sheet or workbook`;
    default:
      return `LET name ${position} (${text}) is not a valid name`;
  }
}

function formatCount(count: number): string {
  return count > 1e9 ? "over a billion" : count.toLocaleString("en-US");
}
