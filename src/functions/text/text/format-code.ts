/**
 * Parser for Excel number format codes, the `format_text` of TEXT.
 *
 * A format code has up to four sections separated by `;`. Each section is a
 * list of codes: digit placeholders (`0 # ?`), a decimal point, grouping and
 * scaling commas, `%`, an exponent (`E+`), a fraction slash, the text
 * placeholder `@`, `General`, date and time codes, literal text and bracketed
 * conditions, colors and locale tags. The parser turns each section into
 * {@link FormatPart}s that the formatters in `format-number.ts` and
 * `format-date.ts` render, and decides which section formats which value
 * (see {@link parseFormatCode}).
 *
 * The rules follow Excel, as reverse-engineered by the numfmt and SSF
 * libraries. Formats that Excel rejects, such as a date code next to a digit
 * placeholder, are reported as invalid.
 */

type ConditionOperator = "<" | "<=" | ">" | ">=" | "=" | "<>";

export type FormatCondition = {
  operator: ConditionOperator;
  operand: number;
};

/** A date or time code. */
export type DateCode =
  /** `yy` or `yyyy` (and `e`): the year with 2 or 4 digits */
  | { type: "year"; digits: 2 | 4 }
  /** `bb` or `bbbb`: the Buddhist year (Gregorian year + 543) */
  | { type: "buddhist-year"; digits: 2 | 4 }
  /** `m` or `mm`: the month number */
  | { type: "month"; pad: boolean }
  /** `mmm`, `mmmm` or `mmmmm` */
  | { type: "month-name"; style: "short" | "long" | "letter" }
  /** `d` or `dd` */
  | { type: "day"; pad: boolean }
  /** `ddd` or `dddd` (also `aaa` and `aaaa`) */
  | { type: "weekday"; style: "short" | "long" }
  /** `h` or `hh` */
  | { type: "hour"; pad: boolean }
  /** `m` or `mm` after an hour or before a second */
  | { type: "minute"; pad: boolean }
  /** `s` or `ss` */
  | { type: "second"; pad: boolean }
  /** `.0`, `.00` or `.000` after a time code: fractions of a second */
  | { type: "subsecond"; digits: number }
  /** `AM/PM` (`letters` undefined) or `A/P` (the letters as written) */
  | { type: "ampm"; letters: [string, string] | undefined }
  /** `[h]`, `[m]` or `[s]`: elapsed hours, minutes or seconds */
  | { type: "elapsed"; unit: "hour" | "minute" | "second"; width: number };

/**
 * Where a literal stands next to a fraction, which decides whether it is
 * shown when the fraction is not (see `format-number.ts`).
 */
type FractionRule = "before-numerator" | "before-slash" | "before-denominator";

/** A run of consecutive digit placeholders, such as the `000` of `000-00`. */
type DigitRun = { pattern: string };

/** One piece of a section's output. */
export type FormatPart =
  /**
   * Literal text: quoted text, an escaped or plain character, a currency
   * symbol or the space of `_x`. `rule` is set for literals next to a
   * fraction; `space` marks a plain space, which fractions treat specially;
   * `fixed` literals (the space of `_x`) are always shown.
   */
  | {
      type: "literal";
      text: string;
      space?: boolean;
      fixed?: boolean;
      rule?: FractionRule;
    }
  | { type: "integer"; run: DigitRun }
  | { type: "point" }
  | { type: "decimals"; run: DigitRun }
  | { type: "exponent"; letter: string }
  | { type: "exponent-digits"; run: DigitRun }
  | { type: "numerator"; run: DigitRun }
  | { type: "slash" }
  | { type: "denominator"; run: DigitRun }
  | { type: "general" }
  | { type: "text" }
  | { type: "date"; code: DateCode }
  /** `*x` repeats x to fill a cell; TEXT has no cell, so it shows nothing */
  | { type: "fill" };

/** A vulgar fraction such as the `?/?` of `# ?/?`. */
export type FractionLayout = {
  /** The numerator placeholders, such as `?` */
  numeratorPattern: string;
  /** The first denominator run, such as `??` or `8` */
  denominatorPattern: string;
  /** A denominator written as digits, such as the 8 of `?/8` */
  fixedDenominator: number | undefined;
  /** The largest denominator that fits the placeholders: 9, 99, ... */
  maxDenominator: number;
  /** Placeholders that show a digit or a space: not `#` */
  numeratorMinimum: number;
  denominatorMinimum: number;
  /** Whether the numerator or denominator contains a `0` */
  zeroPadded: boolean;
};

export type FormatSection = {
  parts: FormatPart[];
  condition: FormatCondition | undefined;
  /**
   * Whether a negative value gets a minus sign: in the single section of a
   * format that has no negative section, and in conditional sections.
   */
  automaticMinus: boolean;
  /** The section contains `General` */
  general: boolean;
  /** The section contains `@` */
  text: boolean;
  /** The section contains date or time codes */
  date: DateLayout | undefined;
  /** All integer placeholders, in order */
  integerPattern: string;
  /** How many runs the integer placeholders are split into */
  integerRuns: number;
  /** All placeholders after the decimal point */
  decimalsPattern: string;
  /** All exponent placeholders */
  exponentPattern: string;
  /** Whether the section has a decimal point */
  point: boolean;
  /** Whether `E+` (rather than `E-`) shows the sign of positive exponents */
  exponentPlus: boolean | undefined;
  fraction: FractionLayout | undefined;
  /** Thousands separators between integer placeholders */
  grouping: boolean;
  /** Power of ten the value is scaled by: +2 per `%`, -3 per scaling comma */
  scale: number;
};

type DateLayout = {
  /** Shows an hour, minute, second or elapsed time, not only a date */
  time: boolean;
  /** Decimals of the fractions of a second shown (0 to 3) */
  secondDecimals: number;
  /** `AM/PM` or `A/P` selects the 12-hour clock */
  twelveHour: boolean;
};

export type ParsedFormatCode = {
  /**
   * The sections for numbers, tried in order: the first whose condition
   * matches (a section without a condition always matches) formats the
   * number. There are always at least two.
   */
  numberSections: FormatSection[];
  /** The section for text, if the format has one */
  textSection: FormatSection | undefined;
};

type FormatCodeError = { error: string };

// ---------------------------------------------------------------------------
// Tokenizer

type Token =
  | { type: "general" }
  | { type: "placeholder"; char: "0" | "#" | "?" }
  | { type: "digit"; char: string }
  | { type: "point" }
  | { type: "group" }
  | { type: "scale"; count: number }
  | { type: "percent" }
  | { type: "exponent"; letter: string; plus: boolean }
  | { type: "slash" }
  | { type: "at" }
  | { type: "literal"; text: string }
  | { type: "space" }
  | { type: "skip" }
  | { type: "fill" }
  | { type: "date"; letters: string }
  | { type: "ampm"; letters: [string, string] | undefined }
  | { type: "elapsed"; unit: "hour" | "minute" | "second"; width: number }
  | { type: "condition"; condition: FormatCondition }
  /** `[$-F800]` or `[$-F400]`: the section shows the system date or time */
  | { type: "system"; pattern: string }
  | { type: "ignored" }
  | { type: "break" };

/** A scaling comma that turns into a thousands separator if digits follow. */
type CommaToken =
  | Extract<Token, { type: "scale" }>
  | Extract<Token, { type: "group" }>;

const isPlaceholderChar = (char: string | undefined) =>
  char === "0" || char === "#" || char === "?";

/** Digits and placeholders: what a comma looks at to decide what it is. */
const isNumberChar = (char: string | undefined) =>
  char !== undefined &&
  (isPlaceholderChar(char) || (char >= "1" && char <= "9"));

const COLOR_PATTERN =
  /^(?:black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)$/i;
const CONDITION_PATTERN = /^(<=|<>|<|>=|>|=)\s*(-?[.\d]+)$/;
const ELAPSED_PATTERN = /^(h+|m+|s+)$/i;
/** Runs of a date or time letter; a lowercase `e` is the year. */
const DATE_PATTERN =
  /^(?:[hH]+|[mM]+|[sS]+|[yY]+|[bB]+|[dD]+|[gG]+|[aA]{3,}|e+)/;
/** A locale tag with a Windows locale id, such as `[$€-407]` or `[$-F800]`. */
const LOCALE_ID_PATTERN = /^-([0-9a-f]{3,8})$/i;
/**
 * The en-US system long date and time formats, which `[$-F800]` and
 * `[$-F400]` select in place of the rest of their section.
 */
const SYSTEM_DATE_FORMAT = "dddd, mmmm d, yyyy";
const SYSTEM_TIME_FORMAT = "h:mm:ss AM/PM";

/**
 * Characters Excel rejects outside quotes: `E` without a sign, `N`, and
 * accented forms of both.
 */
const INVALID_CHARACTERS = new Set("EÈÉÊËèéêëĒēĔĕĖėĘęĚěȄȅȆȇȨȩNnÑñŃńŅņŇňǸǹ");

/** The code point starting at `index`, as a string (one or two UTF-16 units). */
function charAt(code: string, index: number): string | undefined {
  const point = code.codePointAt(index);
  return point === undefined ? undefined : String.fromCodePoint(point);
}

type Tokens = {
  tokens: Token[];
  /**
   * For each section, whether its source starts with a bracket other than a
   * locale tag: `[Red]` or `[>5]` alone format like General.
   */
  bracketStarts: boolean[];
};

function tokenize(code: string): Tokens | FormatCodeError {
  const tokens: Token[] = [];
  const startsWithBracket = (index: number) =>
    code[index] === "[" && code[index + 1] !== "$";
  const bracketStarts = [startsWithBracket(0)];
  // Commas that are scaling unless a digit placeholder follows in the section
  let pendingCommas: CommaToken[] = [];
  // For the exponent `e+`, which otherwise is the year code `e`
  let sectionHasPlaceholder = false;
  let sectionHasDate = false;

  const push = (token: Token) => {
    tokens.push(token);
    if (token.type === "placeholder" || token.type === "digit") {
      for (const comma of pendingCommas) {
        (comma as { type: string }).type = "group";
      }
      pendingCommas = [];
    }
    if (token.type === "placeholder") {
      sectionHasPlaceholder = true;
    }
    if (
      token.type === "date" ||
      token.type === "ampm" ||
      token.type === "elapsed"
    ) {
      sectionHasDate = true;
    }
    if (token.type === "break") {
      pendingCommas = [];
      sectionHasPlaceholder = false;
      sectionHasDate = false;
    }
  };

  let index = 0;
  while (index < code.length) {
    const char = code[index]!;
    const rest = code.slice(index);

    // A comma groups thousands between digits, scales by 1000 after digits
    // or a decimal point, and is a literal comma anywhere else.
    if (char === ",") {
      let count = 1;
      while (code[index + count] === ",") {
        count++;
      }
      const before = code[index - 1];
      const after = code[index + count];
      let group = isNumberChar(before);
      let scale = group || before === ".";
      if (group && (after === undefined || after === ";")) {
        group = false;
      }
      if (scale && isNumberChar(after)) {
        scale = false;
      }
      if (group && !scale) {
        push({ type: "group" });
      } else if (scale) {
        const token: CommaToken = { type: "scale", count };
        push(token);
        if (group) {
          pendingCommas.push(token);
        }
      } else {
        push({ type: "literal", text: "," });
      }
      index += count;
      continue;
    }

    if (/^general/i.test(rest)) {
      push({ type: "general" });
      index += 7;
      continue;
    }

    switch (char) {
      case "0":
      case "#":
      case "?":
        push({ type: "placeholder", char });
        index++;
        continue;
      case "/":
        push({ type: "slash" });
        index++;
        continue;
      case ";":
        push({ type: "break" });
        index++;
        bracketStarts.push(startsWithBracket(index));
        continue;
      case "@":
        push({ type: "at" });
        index++;
        continue;
      case ".":
        push({ type: "point" });
        index++;
        continue;
      case " ":
        push({ type: "space" });
        index++;
        continue;
      case "%":
        push({ type: "percent" });
        index++;
        continue;
      case '"': {
        const end = code.indexOf('"', index + 1);
        if (end < 0) {
          return { error: "a quoted text is not closed" };
        }
        push({ type: "literal", text: code.slice(index + 1, end) });
        index = end + 1;
        continue;
      }
      case "\\":
      case "_":
      case "*": {
        // `_\x` and `*\x` take the escaped character
        let next = index + 1;
        if (char !== "\\" && code[next] === "\\" && next + 1 < code.length) {
          next++;
        }
        const target = charAt(code, next);
        if (target === undefined) {
          return { error: `"${char}" must be followed by a character` };
        }
        if (char === "\\") {
          push({ type: "literal", text: target });
        } else {
          push({ type: char === "_" ? "skip" : "fill" });
        }
        index = next + target.length;
        continue;
      }
      case "[": {
        const end = code.indexOf("]", index);
        if (end < 0) {
          return { error: "a [ is not closed" };
        }
        const token = parseBracket(code.slice(index + 1, end));
        if ("error" in token) {
          return token;
        }
        push(token);
        index = end + 1;
        continue;
      }
    }

    const ampm = /^(am\/pm|a\/p)/i.exec(rest);
    if (ampm) {
      const letters = ampm[1]!;
      push({
        type: "ampm",
        letters: letters.length === 3 ? [letters[0]!, letters[2]!] : undefined,
      });
      index += letters.length;
      continue;
    }

    if (char >= "1" && char <= "9") {
      push({ type: "digit", char });
      index++;
      continue;
    }

    // E+ and E- are exponents. So are e+ and e-, as Microsoft documents them,
    // after digit placeholders; elsewhere e is the year.
    if (
      (char === "E" ||
        (char === "e" && sectionHasPlaceholder && !sectionHasDate)) &&
      (code[index + 1] === "+" || code[index + 1] === "-")
    ) {
      push({ type: "exponent", letter: char, plus: code[index + 1] === "+" });
      index += 2;
      continue;
    }

    // B1 selects the Gregorian calendar (the default), B2 the Hijri calendar
    const calendar = /^b([12])/i.exec(rest);
    if (calendar) {
      if (calendar[1] === "2") {
        return { error: "the Hijri calendar (B2) is not supported" };
      }
      push({ type: "ignored" });
      index += 2;
      continue;
    }
    if (char === "B" && index === code.length - 1) {
      return { error: "a format cannot end with B" };
    }

    const date = DATE_PATTERN.exec(rest);
    if (date) {
      push({ type: "date", letters: date[0] });
      index += date[0].length;
      continue;
    }

    const literal = charAt(code, index)!;
    if (INVALID_CHARACTERS.has(literal)) {
      return { error: `"${literal}" must be quoted` };
    }
    push({ type: "literal", text: literal });
    index += literal.length;
  }
  return { tokens, bracketStarts };
}

/** Parses the inside of `[...]`. */
function parseBracket(content: string): Token | FormatCodeError {
  const elapsed = ELAPSED_PATTERN.exec(content);
  if (elapsed) {
    const letters = elapsed[1]!.toLowerCase();
    return {
      type: "elapsed",
      unit:
        letters[0] === "h" ? "hour" : letters[0] === "m" ? "minute" : "second",
      width: letters.length,
    };
  }
  const condition = CONDITION_PATTERN.exec(content);
  if (condition) {
    const operand = Number(condition[2]);
    if (Number.isNaN(operand)) {
      return { error: `[${content}] is not a valid condition` };
    }
    return {
      type: "condition",
      condition: {
        operator: condition[1] as ConditionOperator,
        operand,
      },
    };
  }
  if (content.startsWith("$")) {
    // [$€-407]: a currency symbol, then a locale after the dash. Only the
    // currency is shown, except for the system date and time locales.
    const dash = content.indexOf("-");
    const currency = dash < 0 ? content.slice(1) : content.slice(1, dash);
    const locale = dash < 0 ? "" : content.slice(dash);
    const id = LOCALE_ID_PATTERN.exec(locale);
    const lcid = id ? parseInt(id[1]!, 16) & 0xffff : undefined;
    if (lcid === 0xf800 || /^-x-sysdate$/i.test(locale)) {
      return { type: "system", pattern: SYSTEM_DATE_FORMAT };
    }
    if (lcid === 0xf400 || /^-x-systime$/i.test(locale)) {
      return { type: "system", pattern: SYSTEM_TIME_FORMAT };
    }
    return currency ? { type: "literal", text: currency } : { type: "ignored" };
  }
  if (COLOR_PATTERN.test(content)) {
    return { type: "ignored" };
  }
  if (/^(?:dbnum|natnum)\d+$/i.test(content)) {
    return { error: `[${content}] is not supported` };
  }
  return { error: `[${content}] is not a valid format code` };
}

// ---------------------------------------------------------------------------
// Sections

type Mode = "integer" | "decimals" | "exponent" | "denominator";

/** A date code while parsing, for telling minutes from months. */
type PendingDateCode = {
  part: Extract<FormatPart, { type: "date" }>;
  /** An hour or second whose minute was found */
  used?: boolean;
  /** An `m` or `mm` read as a month that a following second makes a minute */
  maybeMinute?: boolean;
};

const isHourOrSecond = (code: DateCode) =>
  code.type === "hour" ||
  code.type === "second" ||
  (code.type === "elapsed" && code.unit !== "minute");

const isMinute = (code: DateCode) =>
  code.type === "minute" || (code.type === "elapsed" && code.unit === "minute");

function readDateCode(
  letters: string,
  pending: PendingDateCode[]
): DateCode | undefined {
  const value = letters.toLowerCase();
  const first = value[0];
  const last = pending[pending.length - 1];
  switch (first) {
    case "y":
      return { type: "year", digits: value.length <= 2 ? 2 : 4 };
    case "e":
      return { type: "year", digits: 4 };
    case "b":
      return { type: "buddhist-year", digits: value.length <= 2 ? 2 : 4 };
    case "d":
      return value.length <= 2
        ? { type: "day", pad: value.length === 2 }
        : { type: "weekday", style: value.length === 3 ? "short" : "long" };
    case "a":
      return { type: "weekday", style: value.length === 3 ? "short" : "long" };
    case "h":
      return { type: "hour", pad: value.length > 1 };
    case "s":
      // `m` or `mm` right before seconds are minutes
      if (last?.maybeMinute && last.part.code.type === "month") {
        last.part.code = { type: "minute", pad: last.part.code.pad };
        last.maybeMinute = false;
      }
      return { type: "second", pad: value.length > 1 };
    case "m": {
      if (value.length === 3) {
        return { type: "month-name", style: "short" };
      }
      if (value.length === 5) {
        return { type: "month-name", style: "letter" };
      }
      if (value.length >= 4) {
        return { type: "month-name", style: "long" };
      }
      if (last && !last.used && isHourOrSecond(last.part.code)) {
        last.used = true;
        return { type: "minute", pad: value.length > 1 };
      }
      return { type: "month", pad: value.length > 1 };
    }
    default:
      // g, gg, ggg: the era, which Gregorian dates in English do not show
      return undefined;
  }
}

function parseSection(
  tokens: Token[],
  startsWithBracket: boolean
): FormatSection | FormatCodeError {
  // A system date or time replaces the codes of its section
  const system = tokens.find(
    (token): token is Extract<Token, { type: "system" }> =>
      token.type === "system"
  );
  if (system) {
    const systemTokens = tokenize(system.pattern);
    if ("error" in systemTokens) {
      return systemTokens;
    }
    return parseSection(
      [
        ...tokens.filter((token) => token.type === "condition"),
        ...systemTokens.tokens,
      ],
      false
    );
  }
  const parts: FormatPart[] = [];
  const runs: Record<Mode | "numerator", DigitRun[]> = {
    integer: [],
    decimals: [],
    exponent: [],
    denominator: [],
    numerator: [],
  };
  const pendingDates: PendingDateCode[] = [];
  let mode: Mode = "integer";
  let previous: Token | undefined;
  let lastRunPart: Extract<FormatPart, { run: DigitRun }> | undefined;
  let condition: FormatCondition | undefined;
  let general = false;
  let text = false;
  let isDate = false;
  let time = false;
  let twelveHour = false;
  let secondDecimals = 0;
  let point = false;
  let exponentPlus: boolean | undefined;
  let fraction = false;
  let looseSlash = false;
  let grouping = false;
  let scale = 0;

  const isRunToken = (token: Token | undefined) =>
    token !== undefined &&
    (token.type === "placeholder" ||
      (token.type === "digit" && mode === "denominator"));

  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index]!;
    switch (token.type) {
      case "general":
        general = true;
        parts.push({ type: "general" });
        break;
      case "placeholder":
      case "digit": {
        if (token.type === "digit" && mode !== "denominator") {
          parts.push({ type: "literal", text: token.char });
          break;
        }
        const modeRuns = runs[mode];
        const lastRun = modeRuns[modeRuns.length - 1];
        if (lastRun && (isRunToken(previous) || previous?.type === "group")) {
          lastRun.pattern += token.char;
        } else {
          const run: DigitRun = { pattern: token.char };
          modeRuns.push(run);
          const partType = mode === "exponent" ? "exponent-digits" : mode;
          lastRunPart = { type: partType, run } as Extract<
            FormatPart,
            { run: DigitRun }
          >;
          parts.push(lastRunPart);
        }
        break;
      }
      case "point": {
        // In a date, a point followed by zeros shows fractions of a second
        if (isDate) {
          let digits = 0;
          while (digits < 3) {
            const next = tokens[index + 1];
            if (next?.type !== "placeholder" || next.char !== "0") {
              break;
            }
            digits++;
            index++;
          }
          if (digits > 0) {
            time = true;
            secondDecimals = Math.max(secondDecimals, digits);
            parts.push({ type: "date", code: { type: "subsecond", digits } });
            break;
          }
          parts.push({ type: "literal", text: "." });
          break;
        }
        point = true;
        mode = "decimals";
        parts.push({ type: "point" });
        break;
      }
      case "exponent":
        exponentPlus = token.plus;
        mode = "exponent";
        parts.push({ type: "exponent", letter: token.letter });
        break;
      case "slash": {
        const modeRuns = runs[mode];
        if (modeRuns.length > 0 && lastRunPart) {
          // The digits just before the slash are the numerator
          fraction = true;
          runs.numerator.push(modeRuns.pop()!);
          (lastRunPart as { type: string }).type = "numerator";
          mode = "denominator";
          parts.push({ type: "slash" });
        } else {
          looseSlash = true;
          parts.push({ type: "literal", text: "/" });
        }
        break;
      }
      case "group":
        if (mode === "integer") {
          grouping = true;
        } else if (mode === "denominator") {
          return { error: "a denominator cannot have thousands separators" };
        }
        break;
      case "scale":
        scale -= 3 * token.count;
        break;
      case "percent":
        scale += 2;
        parts.push({ type: "literal", text: "%" });
        break;
      case "at":
        text = true;
        parts.push({ type: "text" });
        break;
      case "literal":
        parts.push({ type: "literal", text: token.text });
        break;
      case "space":
        parts.push({ type: "literal", text: " ", space: true });
        break;
      case "skip":
        parts.push({ type: "literal", text: " ", fixed: true });
        break;
      case "fill":
        parts.push({ type: "fill" });
        break;
      case "date": {
        const code = readDateCode(token.letters, pendingDates);
        if (code) {
          isDate = true;
          if (
            code.type === "hour" ||
            code.type === "minute" ||
            code.type === "second"
          ) {
            time = true;
          }
          const part: Extract<FormatPart, { type: "date" }> = {
            type: "date",
            code,
          };
          parts.push(part);
          const pending: PendingDateCode = { part };
          if (code.type === "month") {
            pending.maybeMinute = true;
          }
          if (code.type === "second") {
            const last = pendingDates[pendingDates.length - 1];
            if (last && isMinute(last.part.code)) {
              pending.used = true;
            }
          }
          pendingDates.push(pending);
        }
        break;
      }
      case "ampm":
        isDate = true;
        time = true;
        twelveHour = true;
        parts.push({
          type: "date",
          code: { type: "ampm", letters: token.letters },
        });
        break;
      case "elapsed": {
        isDate = true;
        time = true;
        const part: Extract<FormatPart, { type: "date" }> = {
          type: "date",
          code: { type: "elapsed", unit: token.unit, width: token.width },
        };
        parts.push(part);
        pendingDates.push({ part });
        break;
      }
      case "condition":
        condition = token.condition;
        break;
      case "system":
      case "ignored":
      case "break":
        break;
    }
    previous = token;
  }

  const hasDigits =
    runs.integer.length +
      runs.decimals.length +
      runs.exponent.length +
      runs.numerator.length +
      runs.denominator.length >
    0;

  if (
    isDate &&
    (hasDigits || scale !== 0 || text || general || exponentPlus !== undefined)
  ) {
    return { error: "date and time codes cannot be mixed with number codes" };
  }
  if (general && (hasDigits || scale !== 0 || text)) {
    return { error: "General cannot be mixed with other number codes" };
  }
  if (text && (hasDigits || scale !== 0 || exponentPlus !== undefined)) {
    return { error: "@ cannot be mixed with number codes" };
  }
  if (fraction && (point || exponentPlus !== undefined)) {
    return { error: "a fraction cannot have a decimal point or an exponent" };
  }
  if (fraction && runs.denominator.length === 0) {
    return { error: "a fraction needs a denominator" };
  }
  if (looseSlash && !fraction && !isDate) {
    return { error: "/ must stand between the digits of a fraction" };
  }
  if (grouping && runs.integer.length === 0) {
    return { error: "a thousands separator needs digit placeholders" };
  }
  if (
    exponentPlus !== undefined &&
    (runs.integer.length + runs.decimals.length === 0 ||
      runs.exponent.length === 0)
  ) {
    return { error: "an exponent needs digit placeholders on both sides" };
  }

  // A pattern of decimals only, such as `.00`, shows the integer as `#`
  if (
    runs.integer.length === 0 &&
    runs.decimals.length > 0 &&
    exponentPlus === undefined
  ) {
    const run: DigitRun = { pattern: "#" };
    runs.integer.push(run);
    const pointIndex = parts.findIndex((part) => part.type === "point");
    parts.splice(pointIndex, 0, { type: "integer", run });
  }

  // `[Red]` or `[>5]` alone formats like General
  if (parts.length === 0 && startsWithBracket) {
    general = true;
    parts.push({ type: "general" });
  }

  // Literals next to a fraction follow its visibility
  if (fraction) {
    for (let index = 0; index < parts.length - 1; index++) {
      const part = parts[index]!;
      if (part.type !== "literal" || part.fixed) {
        continue;
      }
      const next = parts[index + 1]!.type;
      if (next === "numerator") {
        part.rule = "before-numerator";
      } else if (next === "slash") {
        part.rule = "before-slash";
      } else if (next === "denominator") {
        part.rule = "before-denominator";
      }
    }
  }

  const joinRuns = (list: DigitRun[]) =>
    list.map((run) => run.pattern).join("");
  return {
    parts,
    condition,
    automaticMinus: false,
    general,
    text,
    date: isDate ? { time, secondDecimals, twelveHour } : undefined,
    integerPattern: joinRuns(runs.integer),
    integerRuns: runs.integer.length,
    decimalsPattern: joinRuns(runs.decimals),
    exponentPattern: joinRuns(runs.exponent),
    point,
    exponentPlus,
    fraction: fraction
      ? readFractionLayout(runs.numerator, runs.denominator)
      : undefined,
    // Thousands separators only apply to an integer in one piece
    grouping: grouping && runs.integer.length === 1,
    scale,
  };
}

const countNonHash = (pattern: string) =>
  [...pattern].filter((char) => char !== "#").length;

function readFractionLayout(
  numeratorRuns: DigitRun[],
  denominatorRuns: DigitRun[]
): FractionLayout {
  let numeratorPattern = numeratorRuns.map((run) => run.pattern).join("");
  let denominatorPattern = denominatorRuns[0]!.pattern;
  const fixedDigits = denominatorRuns
    .map((run) => run.pattern)
    .join("")
    .replace(/\D/g, "");
  const fixed = fixedDigits === "" ? NaN : Number(fixedDigits);
  // A `?` on either side pads both sides with spaces, also the digits of a
  // fixed denominator
  if (numeratorPattern.includes("?") || denominatorPattern.includes("?")) {
    denominatorPattern = denominatorPattern
      .replace(/\d/g, "?")
      .replace(/#$/, "?");
    numeratorPattern = numeratorPattern.replace(/#$/, "?");
  }
  return {
    numeratorPattern,
    denominatorPattern,
    fixedDenominator: fixed > 0 ? fixed : undefined,
    maxDenominator: 10 ** denominatorPattern.length - 1,
    numeratorMinimum: countNonHash(numeratorPattern),
    denominatorMinimum: countNonHash(denominatorPattern),
    zeroPadded:
      numeratorPattern.includes("0") || denominatorPattern.includes("0"),
  };
}

/**
 * Whether a conditional section shows a minus sign for negative values. A
 * section whose condition only admits negative numbers, such as `[<-5]`,
 * shows the absolute value; any other conditional section shows the sign.
 */
function conditionShowsMinus(condition: FormatCondition | undefined): boolean {
  if (!condition) {
    return true;
  }
  const { operator, operand } = condition;
  const onlyNegative =
    (operand < 0 &&
      (operator === "<" || operator === "<=" || operator === "=")) ||
    (operand === 0 && operator === "<");
  return !onlyNegative;
}

function generalSection(): FormatSection {
  return {
    parts: [{ type: "general" }],
    condition: undefined,
    automaticMinus: false,
    general: true,
    text: false,
    date: undefined,
    integerPattern: "",
    integerRuns: 0,
    decimalsPattern: "",
    exponentPattern: "",
    point: false,
    exponentPlus: undefined,
    fraction: undefined,
    grouping: false,
    scale: 0,
  };
}

/**
 * Resolves which section formats which value, following Excel:
 *
 * - Without conditions the sections are positive; negative; zero; text. A
 *   missing negative section reuses the positive one with a minus sign, and a
 *   missing zero section the positive one. A section containing `@` is the
 *   text section even when there are fewer than four.
 * - With conditions (at most two), the first section whose condition matches
 *   is used; a section without a condition matches any number. A first section
 *   without a condition gets `[>0]`, and a single conditional section is
 *   followed by General. Conditional sections show a minus sign for negative
 *   numbers unless their condition only admits negative numbers. A second
 *   section without a condition shows it only when the first condition is
 *   `[=x]`, `[>x]` or `[>=x]` with x of zero or more.
 */
function resolveSections(
  sections: FormatSection[]
): ParsedFormatCode | FormatCodeError {
  let numberSections: FormatSection[];
  let textSection: FormatSection | undefined;
  const textSections = sections.filter((section) => section.text);
  if (textSections.length > 1) {
    return { error: "only one section can contain @" };
  }
  if (sections.length === 4) {
    textSection = sections[3]!;
    const formatsNumbers =
      textSection.date !== undefined ||
      textSection.fraction !== undefined ||
      textSection.integerRuns > 0 ||
      textSection.decimalsPattern !== "" ||
      textSection.exponentPattern !== "";
    if (formatsNumbers) {
      return {
        error: "the fourth section is for text and cannot format numbers",
      };
    }
    numberSections = sections.slice(0, 3);
    if (numberSections.some((section) => section.text)) {
      return { error: "only the text section can contain @" };
    }
  } else {
    textSection = textSections[0];
    numberSections = sections.filter((section) => !section.text);
  }

  const conditions = numberSections.filter(
    (section) => section.condition
  ).length;
  if (conditions > 2) {
    return { error: "a format can have at most two conditions" };
  }

  if (conditions > 0) {
    const first = numberSections[0]!;
    first.condition ??= { operator: ">", operand: 0 };
    if (numberSections.length === 1) {
      numberSections.push(generalSection());
    }
    if (numberSections.length === 2) {
      const second = numberSections[1]!;
      first.automaticMinus = conditionShowsMinus(first.condition);
      if (second.condition) {
        second.automaticMinus = conditionShowsMinus(second.condition);
      } else {
        const { operator, operand } = first.condition;
        second.automaticMinus =
          operator === "=" ||
          (operand >= 0 && (operator === ">" || operator === ">="));
      }
    } else {
      for (const section of numberSections) {
        section.automaticMinus = conditionShowsMinus(section.condition);
      }
    }
    return { numberSections, textSection };
  }

  if (numberSections.length === 0) {
    numberSections.push(generalSection());
  }
  const positive = numberSections[0]!;
  if (numberSections.length < 2) {
    numberSections.push({ ...positive, automaticMinus: true });
  }
  if (numberSections.length < 3) {
    numberSections.push({ ...positive });
  }
  numberSections[0] = {
    ...numberSections[0]!,
    condition: { operator: ">", operand: 0 },
  };
  numberSections[1] = {
    ...numberSections[1]!,
    condition: { operator: "<", operand: 0 },
  };
  numberSections[2] = { ...numberSections[2]!, condition: undefined };
  return { numberSections, textSection };
}

/** Parsed format codes, most recently added last. */
const cache = new Map<string, ParsedFormatCode | FormatCodeError>();
const MAX_CACHED_FORMATS = 500;

/**
 * Parses a format code, or explains why it is invalid. Results are cached:
 * a workbook typically formats many values with a few codes.
 */
export function parseFormatCode(
  code: string
): ParsedFormatCode | FormatCodeError {
  const cached = cache.get(code);
  if (cached) {
    return cached;
  }
  const parsed = parseUncached(code);
  if (cache.size >= MAX_CACHED_FORMATS) {
    cache.delete(cache.keys().next().value!);
  }
  cache.set(code, parsed);
  return parsed;
}

function parseUncached(code: string): ParsedFormatCode | FormatCodeError {
  const tokenized = tokenize(code);
  if ("error" in tokenized) {
    return tokenized;
  }
  const sectionTokens: Token[][] = [[]];
  for (const token of tokenized.tokens) {
    if (token.type === "break") {
      sectionTokens.push([]);
    } else {
      sectionTokens[sectionTokens.length - 1]!.push(token);
    }
  }
  if (sectionTokens.length > 4) {
    return { error: "a format can have at most four sections" };
  }
  const sections: FormatSection[] = [];
  for (let index = 0; index < sectionTokens.length; index++) {
    const section = parseSection(
      sectionTokens[index]!,
      tokenized.bracketStarts[index] ?? false
    );
    if ("error" in section) {
      return section;
    }
    sections.push(section);
  }
  return resolveSections(sections);
}
