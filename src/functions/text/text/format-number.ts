/**
 * Renders number sections of a format code: digit placeholders, decimals,
 * thousands separators, scaling, exponents, fractions and General.
 *
 * Numbers are formatted as Excel shows them: from their decimal
 * representation with 15 significant digits, rounded half away from zero to
 * the placeholders shown (see `round-utils.ts`). So `0.00` shows 2.675 as
 * "2.68" and 1.005 as "1.01", although their binary values lie just below.
 * Scaling by `%` and trailing commas shifts that decimal representation, so it
 * adds no binary rounding error either.
 */

import { toSignificantDigits } from "../../math/round/round-utils";
import type { FormatPart, FormatSection, FractionLayout } from "./format-code";

/** A non-negative decimal `digits × 10^exponent`; zero has no digits. */
type Decimal = { digits: string; exponent: number };

const ZERO: Decimal = { digits: "", exponent: 0 };

/** The 15-significant-digit decimal of a non-negative finite number. */
function toDecimal(magnitude: number): Decimal {
  if (magnitude === 0) {
    return ZERO;
  }
  const { digits, exponent } = toSignificantDigits(magnitude);
  return normalize({ digits, exponent: exponent - (digits.length - 1) });
}

/** Drops leading and trailing zeros. */
function normalize({ digits, exponent }: Decimal): Decimal {
  let start = 0;
  while (start < digits.length && digits[start] === "0") {
    start++;
  }
  let end = digits.length;
  while (end > start && digits[end - 1] === "0") {
    end--;
  }
  if (start === end) {
    return ZERO;
  }
  return {
    digits: digits.slice(start, end),
    exponent: exponent + (digits.length - end),
  };
}

/** Adds one to a string of decimal digits. */
function increment(digits: string): string {
  let index = digits.length - 1;
  while (index >= 0 && digits[index] === "9") {
    index--;
  }
  if (index < 0) {
    return "1" + "0".repeat(digits.length);
  }
  return (
    digits.slice(0, index) +
    String(Number(digits[index]) + 1) +
    "0".repeat(digits.length - index - 1)
  );
}

/** Rounds to `places` decimals, half away from zero. */
function roundToPlaces(value: Decimal, places: number): Decimal {
  const drop = -value.exponent - places;
  if (drop <= 0 || value.digits === "") {
    return value;
  }
  if (drop > value.digits.length) {
    return ZERO;
  }
  const kept = value.digits.slice(0, value.digits.length - drop);
  const roundsUp = value.digits[value.digits.length - drop]! >= "5";
  return normalize({
    digits: roundsUp ? increment(kept || "0") : kept,
    exponent: -places,
  });
}

/** The decimal exponent of the leading digit: 2 for 123.4, -2 for 0.012. */
const leadingExponent = (value: Decimal) =>
  value.digits.length - 1 + value.exponent;

/**
 * The integer and decimal digits of a decimal: the integer without leading
 * zeros and the decimals without trailing zeros. 12.5 is ["12", "5"], 0.05 is
 * ["", "05"] and 0 is ["", ""].
 */
function splitDigits(value: Decimal): [string, string] {
  const { digits, exponent } = value;
  if (digits === "") {
    return ["", ""];
  }
  if (exponent >= 0) {
    return [digits + "0".repeat(exponent), ""];
  }
  const integerLength = digits.length + exponent;
  return integerLength > 0
    ? [digits.slice(0, integerLength), digits.slice(integerLength)]
    : ["", "0".repeat(-integerLength) + digits];
}

/** What a placeholder shows where there is no digit. */
const pad = (placeholder: string | undefined) =>
  placeholder === "0" ? "0" : placeholder === "?" ? " " : "";

/**
 * Places digits right-aligned across runs of placeholders, such as the runs
 * of `000-00-0000`: the first run takes any digits beyond the placeholders,
 * and runs without digits show their padding.
 */
class RightAlignedRuns {
  private used = 0;

  constructor(
    private readonly digits: string,
    private readonly pattern: string
  ) {}

  next(run: string): string {
    const { digits, pattern } = this;
    const length =
      this.used === 0 && digits.length > pattern.length
        ? run.length + digits.length - pattern.length
        : run.length;
    const offset =
      digits.length < pattern.length
        ? this.used + digits.length - pattern.length
        : this.used;
    let text = "";
    for (let index = 0; index < length; index++) {
      text += digits[index + offset] ?? pad(run[index]);
    }
    this.used += length;
    return text;
  }
}

/**
 * The number of integer positions always shown: up to the leftmost `0` or
 * `?`, counted from the right.
 */
function minimumIntegerLength(pattern: string): number {
  let minimum = 0;
  for (let index = 0; index < pattern.length; index++) {
    const char = pattern[pattern.length - 1 - index];
    if (char !== "#") {
      minimum = index + 1;
    }
  }
  return minimum;
}

/**
 * Renders the integer of a pattern in one piece, such as `#,##0`, with
 * thousands separators when the pattern has them. `zeroFill` shows a zero
 * exponent mantissa with all its placeholders as zeros, as Excel does.
 */
function renderInteger(
  digits: string,
  section: FormatSection,
  zeroFill: boolean
): string {
  const pattern = section.integerPattern;
  const minimum = minimumIntegerLength(pattern);
  const length = zeroFill
    ? Math.max(minimum, pattern.length, digits.length)
    : Math.max(minimum, digits.length);
  let text = "";
  for (let position = length; position > 0; position--) {
    let digit = digits[digits.length - position];
    if (zeroFill && digit === undefined) {
      digit = "0";
    }
    const placeholder =
      digit === undefined
        ? pattern[pattern.length - position] ?? pattern[0]
        : "";
    text += digit ?? pad(placeholder);
    if (section.grouping && position - 1 >= 3 && (position - 1) % 3 === 0) {
      text += digit !== undefined || placeholder === "0" ? "," : " ";
    }
  }
  return text;
}

/**
 * Splits a non-zero decimal into a mantissa rounded to `places` decimals and
 * an exponent, as `0.00E+00` shows 12345 as 1.23 and 4. With several integer
 * placeholders the exponent is a multiple of their count (engineering
 * notation: `##0.0E+0` shows 12345 as 12.3E+3); without any, the mantissa is
 * below 1.
 */
function toScientific(
  value: Decimal,
  places: number,
  integerPlaceholders: number
): { mantissa: Decimal; exponent: number } {
  if (value.digits === "") {
    return { mantissa: ZERO, exponent: 0 };
  }
  const split = (current: Decimal) => {
    const leading = leadingExponent(current);
    let exponent =
      integerPlaceholders > 1
        ? Math.floor(leading / integerPlaceholders) * integerPlaceholders
        : leading;
    if (integerPlaceholders === 0) {
      exponent++;
    }
    const mantissa = roundToPlaces(
      { digits: current.digits, exponent: current.exponent - exponent },
      places
    );
    const rounded = {
      digits: mantissa.digits,
      exponent: mantissa.exponent + exponent,
    };
    return {
      mantissa,
      exponent,
      rounded,
      carried: leadingExponent(rounded) !== leading,
    };
  };
  const first = split(value);
  // Rounding 9.99 up to 10.0 moves the exponent: split the rounded value,
  // a power of ten, again
  return first.carried ? split(first.rounded) : first;
}

type NumberRendering = {
  text: string;
  /**
   * Whether a negative number gets its automatic minus sign: a pattern of
   * digits only shows it when a non-zero digit is shown, so `0` shows -0.4 as
   * "0", as Excel does.
   */
  showsMinus: boolean;
};

/** Renders a number section for the absolute value of a number. */
export function renderNumberSection(
  section: FormatSection,
  magnitude: number
): NumberRendering {
  if (section.fraction) {
    return {
      text: renderFraction(section, section.fraction, magnitude),
      showsMinus: true,
    };
  }

  const decimal = toDecimal(magnitude);
  const value = {
    digits: decimal.digits,
    exponent: decimal.exponent + section.scale,
  };
  const places = section.decimalsPattern.length;
  let integerDigits: string;
  let decimalDigits: string;
  let exponent = 0;
  if (section.exponentPlus !== undefined) {
    const scientific = toScientific(
      value,
      places,
      section.integerPattern.length
    );
    exponent = scientific.exponent;
    [integerDigits, decimalDigits] = splitDigits(scientific.mantissa);
  } else {
    [integerDigits, decimalDigits] = splitDigits(roundToPlaces(value, places));
  }

  const zeroFill =
    section.exponentPlus !== undefined &&
    integerDigits === "" &&
    decimalDigits === "";
  const integerRuns = new RightAlignedRuns(
    integerDigits,
    section.integerPattern
  );
  const exponentRuns = new RightAlignedRuns(
    String(Math.abs(exponent)),
    section.exponentPattern
  );
  let decimalsUsed = 0;
  let exponentSignShown = false;
  let text = "";
  for (const part of section.parts) {
    switch (part.type) {
      case "integer":
        text +=
          section.integerRuns === 1
            ? renderInteger(integerDigits, section, zeroFill)
            : integerRuns.next(part.run.pattern);
        break;
      case "decimals":
        for (let index = 0; index < part.run.pattern.length; index++) {
          text +=
            decimalDigits[decimalsUsed + index] ?? pad(part.run.pattern[index]);
        }
        decimalsUsed += part.run.pattern.length;
        break;
      case "point":
        text += ".";
        break;
      case "exponent":
        text += part.letter;
        break;
      case "exponent-digits":
        if (!exponentSignShown) {
          exponentSignShown = true;
          text += exponent < 0 ? "-" : section.exponentPlus ? "+" : "";
        }
        text += exponentRuns.next(part.run.pattern);
        break;
      case "general":
      case "text":
        // `@` in a number section shows the number as General
        text += formatGeneral(magnitude);
        break;
      default:
        text += renderLiteral(part);
    }
  }

  // Digits decide the minus sign only in patterns that show digits: the
  // integer when there are integer placeholders, the decimals after a point
  const showsDigits = section.integerRuns > 0 || section.point;
  const shownDigits =
    (section.integerRuns > 0 ? integerDigits : "") +
    (section.point ? decimalDigits : "");
  return { text, showsMinus: !showsDigits || shownDigits !== "" };
}

function renderLiteral(part: FormatPart): string {
  return part.type === "literal" ? part.text : "";
}

// ---------------------------------------------------------------------------
// Fractions

/** Denominators beyond this are not distinguishable in a double. */
const MAX_DENOMINATOR = 1e15;

/**
 * The fraction closest to `value` among the convergents of its continued
 * fraction whose denominator is at most `maxDenominator`, stopping once one is
 * within 1e-13. This is Excel's method, applied to the binary value: 0.3 is
 * slightly below 3/10, so `?/?` shows it as 2/7, while 1.3 is slightly above
 * 13/10 and shows as 4/3.
 */
function approximateFraction(
  value: number,
  maxDenominator: number
): [number, number] {
  const limit = Math.min(maxDenominator, MAX_DENOMINATOR);
  let remainder = value;
  let previousNumerator = 0;
  let previousDenominator = 1;
  let numerator = 1;
  let denominator = 0;
  // Denominators grow at least like Fibonacci numbers, so the limit stops
  // the expansion after at most about 75 terms
  for (let term = 0; term < 100; term++) {
    const whole = Math.floor(remainder);
    const nextNumerator = whole * numerator + previousNumerator;
    const nextDenominator = whole * denominator + previousDenominator;
    if (nextDenominator > limit) {
      break;
    }
    previousNumerator = numerator;
    previousDenominator = denominator;
    numerator = nextNumerator;
    denominator = nextDenominator;
    const fraction = remainder - whole;
    if (fraction === 0 || Math.abs(value - numerator / denominator) < 1e-13) {
      break;
    }
    remainder = 1 / fraction;
  }
  return [numerator, denominator];
}

/** Writes a whole number, with 15 significant digits like Excel. */
function wholeNumberText(value: number): string {
  return value < 1e15 ? String(value) : plainText(toDecimal(value));
}

/**
 * Renders a fraction section such as `# ?/?`, `# ??/100` or `?/8`. A mixed
 * fraction (one with integer placeholders) shows the whole number and the
 * remaining fraction, and a fraction that rounds to 0 or to 1 shows as
 * padding instead. Placeholders follow the digit rules: `?` pads with spaces,
 * `0` with zeros and `#` with nothing.
 */
function renderFraction(
  section: FormatSection,
  layout: FractionLayout,
  magnitude: number
): string {
  // Fractions round the binary value, as Excel does: `# ??/100` shows 36.305
  // as 36 30/100
  const value = magnitude * 10 ** section.scale;
  const mixed = section.integerRuns > 0;
  let whole = 0;
  let numerator: number;
  let denominator: number;
  if (layout.fixedDenominator !== undefined) {
    denominator = layout.fixedDenominator;
    if (mixed) {
      whole = Math.floor(value);
      numerator = Math.round((value - whole) * denominator);
    } else {
      numerator = Math.round(value * denominator);
    }
  } else if (mixed) {
    whole = Math.floor(value);
    [numerator, denominator] = approximateFraction(
      value - whole,
      layout.maxDenominator
    );
  } else {
    [numerator, denominator] = approximateFraction(
      value,
      layout.maxDenominator
    );
  }
  if (mixed && numerator === denominator && numerator !== 0) {
    whole++;
    numerator = 0;
  }

  // A mixed fraction hides a zero fraction; patterns with a 0 show 0/1 instead
  const hasFraction = !mixed || numerator !== 0;
  const showsFraction = hasFraction || layout.zeroPadded;
  let integerText = whole > 0 ? wholeNumberText(whole) : "";
  if (mixed && !showsFraction && whole === 0) {
    integerText = "0";
  }
  const numeratorText = hasFraction ? wholeNumberText(numerator) : "";
  const denominatorText = hasFraction ? wholeNumberText(denominator) : "";

  const integerRuns = new RightAlignedRuns(integerText, section.integerPattern);
  const numeratorRuns = new RightAlignedRuns(
    numeratorText,
    layout.numeratorPattern
  );
  const { numeratorMinimum, denominatorMinimum } = layout;
  let denominatorUsed = 0;
  let fixedPadding = false;
  let text = "";
  for (const part of section.parts) {
    switch (part.type) {
      case "integer":
        text +=
          section.integerRuns === 1
            ? renderInteger(integerText, section, false)
            : integerRuns.next(part.run.pattern);
        break;
      case "numerator":
        text += numeratorRuns.next(part.run.pattern);
        break;
      case "slash":
        text += showsFraction
          ? "/"
          : numeratorMinimum > 0 || denominatorMinimum > 0
          ? " "
          : "";
        break;
      case "denominator": {
        const pattern = part.run.pattern;
        for (let index = 0; index < pattern.length; index++) {
          const digit = denominatorText[denominatorUsed + index];
          if (digit !== undefined) {
            text += digit;
            continue;
          }
          const char = pattern[index]!;
          if ((char >= "1" && char <= "9") || (fixedPadding && char === "0")) {
            // The digits of a fixed denominator pad with spaces
            fixedPadding = true;
            text += " ";
          } else if (
            !fixedPadding &&
            index === pattern.length - 1 &&
            char === "0" &&
            denominatorText === ""
          ) {
            text += "1";
          } else {
            text += pad(char);
          }
        }
        denominatorUsed += pattern.length;
        break;
      }
      case "literal":
        text += renderFractionLiteral(part, {
          showsFraction,
          hasInteger: integerText !== "",
          numeratorMinimum,
          denominatorMinimum,
        });
        break;
      case "general":
      case "text":
        text += formatGeneral(magnitude);
        break;
      default:
        text += renderLiteral(part);
    }
  }
  return text;
}

/**
 * Literals next to a fraction are shown with it; when the fraction is not
 * shown they become padding, or disappear when its placeholders are all `#`.
 */
function renderFractionLiteral(
  part: Extract<FormatPart, { type: "literal" }>,
  state: {
    showsFraction: boolean;
    hasInteger: boolean;
    numeratorMinimum: number;
    denominatorMinimum: number;
  }
): string {
  const { showsFraction, hasInteger, numeratorMinimum, denominatorMinimum } =
    state;
  const { rule } = part;
  if (!rule) {
    return part.text;
  }
  if (part.space) {
    if (rule !== "before-numerator") {
      return " ";
    }
    return (showsFraction || numeratorMinimum > 0 || denominatorMinimum > 0) &&
      (hasInteger || numeratorMinimum > 0)
      ? " "
      : "";
  }
  if (
    (rule === "before-slash" && showsFraction) ||
    (rule === "before-numerator" && showsFraction && hasInteger) ||
    (rule === "before-denominator" && showsFraction)
  ) {
    return part.text;
  }
  if (
    denominatorMinimum > 0 ||
    (rule === "before-slash" && numeratorMinimum > 0) ||
    (rule === "before-numerator" &&
      denominatorMinimum > 0 &&
      (hasInteger || numeratorMinimum > 0))
  ) {
    return " ".repeat([...part.text].length);
  }
  return "";
}

// ---------------------------------------------------------------------------
// General

/** The widest number General shows, in characters, not counting a minus sign. */
const GENERAL_WIDTH = 11;

/** Writes a decimal in plain notation: "1234.5", "0.00012". */
function plainText(value: Decimal): string {
  const [integer, decimals] = splitDigits(value);
  return (integer || "0") + (decimals ? "." + decimals : "");
}

/**
 * Writes a non-zero decimal in General's scientific notation: a mantissa with
 * up to 5 decimals (4 from an exponent of 99 on, to stay within 11
 * characters) and an exponent of at least two digits, such as "1.23457E+11".
 */
function scientificText(value: Decimal): string {
  const leading = leadingExponent(value);
  const decimals = Math.abs(leading) >= 99 ? 4 : 5;
  const { mantissa, exponent } = toScientific(value, decimals, 1);
  const digits = String(Math.abs(exponent)).padStart(2, "0");
  return `${plainText(mantissa)}E${exponent < 0 ? "-" : "+"}${digits}`;
}

/**
 * Excel's General format for the absolute value of a number, as TEXT and a
 * standard-width cell show it: the number as it is when it fits in 11
 * characters, otherwise rounded to fit (`1/3` is "0.333333333"), and in
 * scientific notation when the integer has more than 11 digits or the number
 * is below 0.0001 and does not fit ("1.23457E+11", "1.23457E-05").
 */
function formatGeneral(magnitude: number): string {
  const value = toDecimal(magnitude);
  if (value.digits === "") {
    return "0";
  }
  const plain = plainText(value);
  if (plain.length <= GENERAL_WIDTH) {
    return plain;
  }
  const leading = leadingExponent(value);
  if (leading >= GENERAL_WIDTH || leading <= -5) {
    return scientificText(value);
  }
  // Round to the decimals that fit: 11 characters less the integer and point
  const decimals = Math.max(0, 9 - Math.max(leading, 0));
  const rounded = plainText(roundToPlaces(value, decimals));
  return rounded.length <= GENERAL_WIDTH ? rounded : scientificText(value);
}
