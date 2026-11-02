import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import type { ASTNode } from "../../../parser/ast";
import { expandLet, LetExpansionError } from "./let-expansion";

/**
 * LET(name1, name_value1, calculation_or_name2, [name_value2, calculation_or_name3], ...)
 *
 * Assigns names to values inside one formula and returns the final
 * calculation, which can use every name. Names are local to the LET and match
 * without regard to case; each value can use the names declared before it;
 * a nested LET can redeclare an outer name; and a LET name hides a defined
 * name (named expression) of the same name for unqualified references.
 *
 * LET is a special form: its arguments are never evaluated as such. The call
 * is expanded by hygienic substitution ({@link expandLet}): each use of a name
 * is replaced with the expression bound to it, and the expanded expression is
 * evaluated in place of the call. Because the expanded expression is ordinary
 * formula AST, everything keyed by formula text keeps working: the AST cache
 * (an environment of evaluated values would make `x*2` collide across cells),
 * dependency tracking, lazily evaluated spills (an array literal like
 * `{x, x+1}` is evaluated cell by cell after LET returns), context-dependent
 * functions such as `ROW()`, recalculation and snapshots. A bound reference
 * stays a reference, so `LET(c, A2, ROW(c))` is 2 and `LET(r, A1:A3, SUM(r))`
 * sums the range.
 *
 * Errors:
 * - `#VALUE!` when the argument count is not odd and at least 3, or the final
 *   calculation is omitted.
 * - `#NAME?` when a name argument is not a plain name (`x1` is a cell
 *   reference; `Sheet1!x` is qualified) or a name is repeated within one LET.
 *   Excel refuses to enter such formulas at all.
 * - `#CALC!` when the expansion exceeds `MAX_LET_EXPANSION_NODES`.
 * An error inside a name value only matters where the name is used.
 *
 * Differences from Excel:
 * - Excel calculates each name value once. Here a name used n times
 *   evaluates its value n times per calculation pass, like writing the
 *   expression n times (the engine does not reuse identical sub-expressions
 *   within a pass). Recalculation of unchanged expressions is still cached.
 * - A name value that is never used is never evaluated.
 * - An omitted name value (`LET(x, , x+1)`) is 0.
 * - `c`, `r`, `C` and `R` are valid names: Excel reserves them for R1C1
 *   notation, which this engine does not support.
 * - Array literals may contain names (`LET(x, 5, {x, x+1})`); Excel only
 *   allows constants in array literals.
 */
export const LET: FunctionDefinition = {
  name: "LET",
  evaluate: function (node, context): FunctionEvaluationResult {
    let expanded: ASTNode;
    try {
      expanded = expandLet(node);
    } catch (error) {
      if (!(error instanceof LetExpansionError)) throw error;
      return {
        type: "error",
        err: error.formulaError,
        message: error.message,
        errAddress: context.dependencyNode,
      };
    }
    return this.evaluateNode(expanded, context);
  },
};
