import {
  FormulaError,
  type CellValue,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../../../core/types";
import {
  createBroadcastResult,
  toBroadcastOperand,
} from "../../../evaluator/broadcast";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { getSuppliedArgument } from "../../../parser/ast";
import { parseFormatCode } from "./format-code";
import { formatNumberWithCode, formatTextWithCode } from "./format-value";

const ZERO: FunctionEvaluationResult = {
  type: "value",
  result: { type: "number", value: 0 },
};

const EMPTY_FORMAT: FunctionEvaluationResult = {
  type: "value",
  result: { type: "string", value: "" },
};

function describeValue(value: CellValue): string {
  return value.type === "string" && value.value === ""
    ? "an empty value"
    : value.type;
}

const infinityText = (negative: boolean) =>
  negative ? "-INFINITY" : "INFINITY";

/**
 * Formats one value with one format code, or explains why it cannot.
 */
function formatValue(
  value: CellValue,
  format: CellValue,
  context: EvaluationContext
): SingleEvaluationResult {
  const error = (
    err: FormulaError,
    message: string
  ): SingleEvaluationResult => ({
    type: "error",
    err,
    message: `TEXT ${message}`,
    errAddress: context.dependencyNode,
  });
  const text = (result: string): SingleEvaluationResult => ({
    type: "value",
    result: { type: "string", value: result },
  });

  if (format.type !== "string") {
    return error(
      FormulaError.VALUE,
      `format_text must be text, got ${describeValue(format)}`
    );
  }
  const parsed = parseFormatCode(format.value);
  if ("error" in parsed) {
    return error(
      FormulaError.VALUE,
      `format_text "${format.value}" is not a valid format: ${parsed.error}`
    );
  }

  switch (value.type) {
    case "number": {
      if (Number.isNaN(value.value)) {
        return error(FormulaError.NUM, "value is not a number");
      }
      if (!Number.isFinite(value.value)) {
        return text(infinityText(value.value < 0));
      }
      const outcome = formatNumberWithCode(value.value, parsed);
      return "error" in outcome
        ? error(
            FormulaError.VALUE,
            `cannot format the value with "${format.value}": ${outcome.error}`
          )
        : text(outcome.text);
    }
    case "infinity":
      return text(infinityText(value.sign === "negative"));
    case "boolean":
      return text(formatTextWithCode(value.value ? "TRUE" : "FALSE", parsed));
    case "string":
      return text(formatTextWithCode(value.value, parsed));
  }
}

/**
 * TEXT(value, format_text)
 *
 * Formats a value as text with an Excel number format code:
 * `TEXT(1234.567, "$#,##0.00")` is "$1,234.57", `TEXT(0.285, "0.0%")` is
 * "28.5%" and `TEXT(45000, "dddd, mmmm d, yyyy")` is
 * "Wednesday, March 15, 2023".
 *
 * Format codes (see `format-code.ts`):
 * - Digit placeholders `0` (shows a zero where there is no digit), `#` (shows
 *   nothing) and `?` (shows a space), before and after a decimal point. A comma
 *   between integer placeholders groups thousands, each trailing comma divides
 *   by 1000 and each `%` multiplies by 100. Digits fill runs of placeholders
 *   separated by literal text from the right, as in `000-00-0000`.
 * - Scientific notation `0.00E+00` (also `E-`, and `e+` and `e-` after digit
 *   placeholders); several integer placeholders, as in `##0.0E+0`, make the
 *   exponent a multiple of their count.
 * - Fractions such as `# ?/?`, `# ??/??`, `?/?` and `# ?/8`: the closest
 *   convergent of the continued fraction, as Excel finds it.
 * - `General` (case-insensitive): Excel's General format, which shows a number
 *   in at most 11 characters, switching to scientific notation for large and
 *   small numbers ("1.23457E+11"). `@` is the text placeholder.
 * - Dates and times on serial numbers of the 1900 date system, including
 *   Excel's 1900-01-00 (serial 0) and fictitious 1900-02-29 (serial 60): `yy`
 *   `yyyy`, `m` `mm` `mmm` `mmmm` `mmmmm`, `d` `dd` `ddd` `dddd`, `h` `hh`,
 *   `m` `mm` as minutes after an hour or before seconds, `s` `ss`, fractions
 *   of a second `ss.00`, `AM/PM` and `A/P` for the 12-hour clock, and elapsed
 *   time `[h]` `[m]` `[s]`. The time is rounded to the seconds shown (or to
 *   their fraction), never to minutes, and a time rounded up to midnight
 *   moves to the next day.
 * - Literal text: `"quoted"`, `\x`, and `$ - + ( ) : ! ^ & ' ~ { } < > =` and
 *   space unquoted. `_x` shows a space and `*x` (fill a cell) shows nothing.
 * - Up to four sections `positive;negative;zero;text`. A negative section
 *   formats the absolute value; without one, a negative number gets a minus
 *   sign, but only when a non-zero digit is shown (`TEXT(-0.4, "0")` is "0",
 *   as in Excel). Up to two conditions such as `[>=100]` choose sections
 *   instead, with Excel's rules for their minus signs.
 * - Colors (`[Red]`, `[Color10]`) and locale tags (`[$-409]`) are ignored, and
 *   a currency tag shows its symbol (`[$€-407]` is "€"). `[$-F800]` and
 *   `[$-F400]` show the en-US system long date and time.
 *
 * Values:
 * - Numbers are formatted from their decimal representation with 15
 *   significant digits, rounded half away from zero: `TEXT(2.675, "0.00")` is
 *   "2.68" and `TEXT(1.005, "0.00")` is "1.01", as in Excel. A number exactly
 *   halfway between two 15-digit decimals is shown as the one nearer zero, as
 *   Excel shows it: `TEXT(1000000000000005, "0")` is "1000000000000000".
 * - Text is returned unchanged, or formatted by the format's text section (the
 *   fourth section, or the section containing `@`).
 * - TRUE and FALSE are formatted as the text "TRUE" and "FALSE".
 * - INFINITY and -INFINITY give "INFINITY" and "-INFINITY" whatever the
 *   format, the text the engine shows for them.
 * - An omitted value (`TEXT(,"0")`) is 0 and an omitted format_text
 *   (`TEXT(5,)`) is the empty format, which shows nothing for numbers.
 * - Errors propagate; an error in `value` wins over one in `format_text`.
 * - #VALUE! for: a format_text that is not text, an invalid format code, a
 *   format without a section for the number (`[>0]0;[<0]0` and 0), a date or
 *   time format for a serial outside 0 to 2958465 (9999-12-31; negative
 *   elapsed times included), and the unsupported codes `[DBNum]`, `[NatNum]`
 *   and `B2` (Hijri calendar). Exactly two arguments are required.
 * - Array arguments are formatted element by element and broadcast against
 *   each other like the operands of an operator (see
 *   `evaluator/broadcast.ts`): `TEXT(A1:A3, "0.0")` spills three values and
 *   `TEXT(5, {"0","0.0"})` two. An error in one element only affects that
 *   element, and a blank cell of a range reads as empty text.
 *
 * Deviations from Excel:
 * - Text is never converted to a number, following this engine's strict
 *   typing: `TEXT("1.5", "0.00")` is "1.5", where Excel shows "1.50".
 * - An empty cell reads as empty text, so `TEXT(A1, "0.00")` of an empty A1
 *   is "", where Excel formats it as 0 ("0.00").
 * - format_text must be text: Excel converts `TEXT(5, 0)` to `TEXT(5, "0")`.
 * - Month and day names are English and the separators `.` and `,` whatever
 *   the locale tag, which in Excel also selects the language and calendar.
 * - `[DBNum]`, `[NatNum]` and `B2` give #VALUE!, where Excel shows East Asian
 *   numerals or Hijri dates.
 */
export const TEXT: FunctionDefinition = {
  name: "TEXT",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "TEXT function takes exactly 2 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate both arguments before returning either, so that one pass
    // registers all dependencies. An omitted value is 0 and an omitted
    // format the empty format, as in Excel.
    const valueNode = getSuppliedArgument(node, 0);
    const formatNode = getSuppliedArgument(node, 1);
    const valueResult = valueNode
      ? this.evaluateNode(valueNode, context)
      : ZERO;
    const formatResult = formatNode
      ? this.evaluateNode(formatNode, context)
      : EMPTY_FORMAT;
    if (
      valueResult.type === "error" ||
      valueResult.type === "awaiting-evaluation"
    ) {
      return valueResult;
    }
    if (
      formatResult.type === "error" ||
      formatResult.type === "awaiting-evaluation"
    ) {
      return formatResult;
    }

    if (valueResult.type === "value" && formatResult.type === "value") {
      return formatValue(valueResult.result, formatResult.result, context);
    }

    return createBroadcastResult({
      operands: [
        toBroadcastOperand(valueResult, context.cellAddress),
        toBroadcastOperand(formatResult, context.cellAddress),
      ],
      context,
      source: "TEXT function",
      combine: (operandCell) => {
        const value = operandCell(0);
        if (value.type !== "value") {
          return value;
        }
        const format = operandCell(1);
        if (format.type !== "value") {
          return format;
        }
        return formatValue(value.result, format.result, context);
      },
    });
  },
};
