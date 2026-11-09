/**
 * Formats a value with a parsed format code: picks the section that applies,
 * renders it and adds the automatic minus sign.
 */

import type {
  FormatCondition,
  FormatSection,
  ParsedFormatCode,
} from "./format-code";
import { renderDateSection } from "./format-date";
import { renderNumberSection } from "./format-number";

type FormatOutcome = { text: string } | { error: string };

function matches(
  condition: FormatCondition | undefined,
  value: number
): boolean {
  if (!condition) {
    return true;
  }
  const { operator, operand } = condition;
  switch (operator) {
    case "<":
      return value < operand;
    case "<=":
      return value <= operand;
    case ">":
      return value > operand;
    case ">=":
      return value >= operand;
    case "=":
      return value === operand;
    case "<>":
      return value !== operand;
  }
}

/** Renders a section for a number; `value` may be negative. */
function renderSection(section: FormatSection, value: number): FormatOutcome {
  if (section.parts.length === 0) {
    return { text: "" };
  }
  if (section.date) {
    const text = renderDateSection(section, value);
    return text === undefined
      ? {
          error: `${value} is not a date or time that can be shown (dates run from 0, 1900-01-00, to 2958465, 9999-12-31)`,
        }
      : { text };
  }
  const { text, showsMinus } = renderNumberSection(section, Math.abs(value));
  const minus = section.automaticMinus && value < 0 && showsMinus;
  return { text: minus ? "-" + text : text };
}

/**
 * Formats a finite number: the first section whose condition matches formats
 * it (see `parseFormatCode` for how sections get their conditions).
 */
export function formatNumberWithCode(
  value: number,
  format: ParsedFormatCode
): FormatOutcome {
  if (!Number.isFinite(value)) {
    return { error: `${value} is not a finite number` };
  }
  const section = format.numberSections.find((candidate) =>
    matches(candidate.condition, value)
  );
  if (!section) {
    return { error: `no section of the format applies to ${value}` };
  }
  return renderSection(section, value);
}

/**
 * Formats text with the text section: `@` is replaced by the text. Without a
 * text section the text is returned unchanged.
 */
export function formatTextWithCode(
  value: string,
  format: ParsedFormatCode
): string {
  const section = format.textSection;
  if (!section) {
    return value;
  }
  let text = "";
  for (const part of section.parts) {
    switch (part.type) {
      case "text":
      case "general":
        text += value;
        break;
      case "literal":
        text += part.text;
        break;
      case "point":
        text += ".";
        break;
      default:
        break;
    }
  }
  return text;
}
