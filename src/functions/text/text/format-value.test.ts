import { describe, expect, test } from "bun:test";
import { parseFormatCode } from "./format-code";
import { formatNumberWithCode, formatTextWithCode } from "./format-value";

/**
 * The formatter behind TEXT, without the engine. Expected results follow
 * Excel; most come from the Excel-generated test tables of the SSF and numfmt
 * libraries.
 */

const ERROR = Symbol("error");

function format(code: string, value: number | string): string | typeof ERROR {
  const parsed = parseFormatCode(code);
  if ("error" in parsed) {
    return ERROR;
  }
  if (typeof value === "string") {
    return formatTextWithCode(value, parsed);
  }
  const outcome = formatNumberWithCode(value, parsed);
  return "error" in outcome ? ERROR : outcome.text;
}

type Case = [
  code: string,
  value: number | string,
  expected: string | typeof ERROR
];

function cases(list: Case[]) {
  test.each(list)("%p formats %p as %p", (code, value, expected) => {
    expect(format(code, value)).toBe(expected);
  });
}

describe("digit placeholders and rounding", () => {
  cases([
    ["0", 1234.5, "1235"],
    ["0", -1234.5, "-1235"],
    ["0", 0.5, "1"],
    ["0", -0.5, "-1"],
    ["0", 2.5, "3"],
    ["0", -2.5, "-3"],
    // 15 significant digits are rounded, not the binary value
    ["0.00", 2.675, "2.68"],
    ["0.00", 1.005, "1.01"],
    ["0.00", 0.125, "0.13"],
    ["0", 2.5 - 2 ** -51, "3"],
    ["0.00000000000000000000", 1 / 3, "0.33333333333333300000"],
    ["0", 123456789012345678, "123456789012346000"],
    ["#,##0", 1e21, "1,000,000,000,000,000,000,000"],
    ["0", 1.5e-300, "0"],
    // A minus sign only when a non-zero digit is shown
    ["0", -0.4, "0"],
    ["0.0", -0.04, "0.0"],
    ["0.0", -0.05, "-0.1"],
    ["0.00", -0.001, "0.00"],
    ["#,##0", 1234567.891, "1,234,568"],
    ["#,##0.00", -1234567.891, "-1,234,567.89"],
    ["0000", 12, "0012"],
    ["#", 0, ""],
    ["#.#", 0, "."],
    ["#.##", 1, "1."],
    ["0.0#", 1.5, "1.5"],
    ["0.0#", 1.25, "1.25"],
    ["0.0#", 0.0012345, "0.0"],
    ["???.??", 1.5, "  1.5 "],
    ["0?", 0, "0 "],
    ["0#", 0, "0"],
    [".00", 0.5, ".50"],
    [".0", 1234, "1234.0"],
    ["#.##0", -10.29, "-10.290"],
    // A second point is shown as written, and the decimals continue after it
    ["0.0.0", 1.25, "1.2.5"],
  ]);
});

describe("thousands separators and scaling", () => {
  cases([
    ["#,##0,", 1234567, "1,235"],
    ["0.0,,", 1234567, "1.2"],
    ["0,", 499, "0"],
    ["0,", 500, "1"],
    ["0,,,", 1234567.89, "0"],
    ["0,0", 1234567.89, "1,234,568"],
    ["0.0,0", 1234567.89, "1234567.89"],
    ['#,##0.0," K"', 12345, "12.3 K"],
    ['#,,"M"', 6000000, "6M"],
    ["0.0%", 0.285, "28.5%"],
    ["0%", 0.285, "29%"],
    // Each % multiplies by 100
    ["0%%", 0.5, "5000%%"],
    ["%0", 0.5, "%50"],
    // A comma that is neither between nor after digits is shown
    ["0 , 0", 1234567.89, "123456 , 8"],
    ["0,x0", 1234567.89, "123456x8"],
    ["x,0", 1234567.89, "x,1234568"],
  ]);
});

describe("literals and digit groups", () => {
  cases([
    ["000-00-0000", 123456789, "123-45-6789"],
    ["000-00-0000", 12345, "000-01-2345"],
    ["(000) 000-0000", 5551234567, "(555) 123-4567"],
    ["###\\###\\##0.00", 101, "#1#01.00"],
    ["##0° 00' 00''", 123456, "12° 34' 56''"],
    ['"Well "00', 5, "Well 05"],
    ['0" µL"', 12, "12 µL"],
    ["\\$0.00", 5, "$5.00"],
    ["$#,##0.00", -1234.5, "-$1,234.50"],
    ["$-+():!^&'~{}<>= 0", 5, "$-+():!^&'~{}<>= 5"],
    ['"a;b" 0', 5, "a;b 5"],
    ["0_)", 5, "5 "],
    ["_(0_)", 5, " 5 "],
    // TEXT has no cell to fill, so *x shows nothing
    ["*-0", 5, "5"],
    ["0*x", 5, "5"],
    ["x", -1, "-x"],
  ]);
});

describe("scientific notation", () => {
  cases([
    ["0.00E+00", 12200000, "1.22E+07"],
    ["0.00E+00", 0.000123, "1.23E-04"],
    ["0.00E-00", 12345, "1.23E04"],
    ["0.00E-00", 0.00012345, "1.23E-04"],
    ["0.00E+00", 0, "0.00E+00"],
    ["0.00E+00", -12345, "-1.23E+04"],
    ["0.00E+00", 9.995, "1.00E+01"],
    ["0E+0", 1e100, "1E+100"],
    ["0.00E+00", 1.7976931348623157e308, "1.80E+308"],
    // Several integer placeholders make the exponent a multiple of their count
    ["##0.0E+0", 12345, "12.3E+3"],
    ["##0.0E+0", 999.96, "1.0E+3"],
    ["##0.0E+0", 0.0012345, "1.2E-3"],
    ["000E+00", 12345.67, "012E+03"],
    [".0E+0", 12345, ".1E+5"],
    // Lowercase e+ and e- work too and show a lowercase e
    ["0.0e+0", 12345, "1.2e+4"],
    ["0.00E+0x0", 123456789, "1.23E+0x8"],
  ]);
});

describe("fractions", () => {
  cases([
    ["# ?/?", 4.34, "4 1/3"],
    ["# ?/?", 0.34, " 1/3"],
    // Continued fractions of the binary value: 0.3 is just below 3/10
    ["# ?/?", 0.3, " 2/7"],
    ["# ?/?", 1.3, "1 1/3"],
    ["# ?/?", 2.3, "2 2/7"],
    ["# ?/?", 5, "5    "],
    ["# ?/?", -1.2, "-1 1/5"],
    ["# ?/?", -0.01, "-0    "],
    // A fraction that rounds to 1 moves into the whole number
    ["# ?/?", 4.96, "5    "],
    ["# ??/??", 12.3456789, "12 28/81"],
    ["# ??/??", -1.2, "-1  1/5 "],
    ["??/??", 12.3456789, "1000/81"],
    ["??/??", -1.2, "- 6/5 "],
    ["?/?", 0, "0/1"],
    ["?/?", 12.345, "37/3"],
    ["# ??/?????????", 0.123251512342345, " 480894/3901729  "],
    ["# ?/8", 0.3, " 2/8"],
    ["# ?/8", 4.97, "5    "],
    // Fixed denominators round the binary value: 0.305 * 100 is 30.49999...
    ["# ??/100", 36.305, "36 30/100"],
    ["# ??/100", 1, "1       "],
    ["0/0", 0, "0/1"],
    ["0 0/0", 1, "1 0/1"],
    ["# #/#", 0, "0"],
    ["# #/#", 12.345, "12 1/3"],
    ["0,0 0/0", 1234567.89, "1,234,567 8/9"],
  ]);
});

describe("General", () => {
  cases([
    ["General", 1 / 3, "0.333333333"],
    ["General", -1 / 3, "-0.333333333"],
    ["General", 2 / 3, "0.666666667"],
    ["General", 123456789012, "1.23457E+11"],
    ["General", 12345678901, "12345678901"],
    ["General", 12345678901.5, "12345678902"],
    ["General", 99999999999.6, "1E+11"],
    ["General", 1234567890.12, "1234567890"],
    ["General", 123456.789012, "123456.789"],
    ["General", 0.0001, "0.0001"],
    ["General", 0.0001234567, "0.000123457"],
    ["General", 0.00001, "0.00001"],
    ["General", 0.00001234567, "1.23457E-05"],
    ["General", 1e-9, "0.000000001"],
    ["General", 1e-10, "1E-10"],
    ["General", 0.1 + 0.2, "0.3"],
    ["General", 1e21, "1E+21"],
    ["General", 1.7976931348623157e308, "1.7977E+308"],
    ["General", 0, "0"],
    ["general", 5, "5"],
    ['"Total: "General', -5, "-Total: 5"],
    ['General" units"', 1.5, "1.5 units"],
    ["General;(General)", -5, "(5)"],
    // A section of only a color or condition formats like General
    ["[Red]", -5, "-5"],
    ["[>5]", 3, "3"],
  ]);
});

describe("sections", () => {
  cases([
    ["0;(0)", -5, "(5)"],
    ["0;(0)", 0, "0"],
    ['0;-0;"zero"', 0, "zero"],
    ['0;-0;"zero"', 0.001, "0"],
    ["0;-0;0", -0, "0"],
    // An explicit negative section shows a rounded-off minus sign
    ["0.0;-0.0", -0.01, "-0.0"],
    ["0;;", -5, ""],
    [";;;", 5, ""],
    ["", 5, ""],
    ["", -5, ""],
    ["0;-0;;@", 0, ""],
    // A section with @ formats text, even when it is not the fourth
    ['0;"t"@', -5, "-5"],
    ['@" units"', 5, "5"],
  ]);
});

describe("conditions", () => {
  cases([
    ["[>=100]0;[<100]0.0", 5, "5.0"],
    ["[>=100]0;[<100]0.0", 500, "500"],
    ["[>=100]0;[<100]0.0", -5, "-5.0"],
    ['[<=-5]"low";0', -10, "low"],
    // Excel shows no minus sign here: the second section follows [<=-5]
    ['[<=-5]"low";0', -1, "1"],
    ["[<1]0.00;0", -3, "-3.00"],
    ['[=1]"one";[=2]"two";"other"', 2, "two"],
    ['[=1]"one";[=2]"two";"other"', -3, "-other"],
    ["[>=100]#,##0;[<=-100]-#,##0;#,##0.00", -3.96, "-3.96"],
    ["[>=100]#,##0;[<=-100]-#,##0;#,##0.00", -1234, "-1,234"],
    ["[<=9999999]###-####;(###) ###-####", 8675309, "867-5309"],
    ["[<=9999999]###-####;(###) ###-####", 2813308004, "(281) 330-8004"],
    ["[>10]0", 5, "5"],
    ["[>10]0", -5, "-5"],
    ["[Red][=50]General;[Blue]000", 49, "049"],
    ["[<-25]General;[>25]General;General", -26, "26"],
    ["[<-25]General;[>25]General;General", -1, "-1"],
    ["[< 5]0.0;0", 3, "3.0"],
    // No section applies
    ['[>0]"pos";[<0]"neg"', 0, ERROR],
  ]);
});

describe("colors and locale tags", () => {
  cases([
    ["[Red]0.00;[Blue]-0.00", -1.5, "-1.50"],
    ["[Color10]0", 5, "5"],
    ["[color 3]0", 5, "5"],
    ["[$-409]0", 5, "5"],
    ["[$-409]", 5, ""],
    ["[$€-407]0.00", 5, "€5.00"],
    ["[$$-409]#,##0.00", 1234.5, "$1,234.50"],
    ["[$INR] #,##0", 1234, "INR 1,234"],
    ["[$]0", 1, "1"],
  ]);
});

describe("text", () => {
  cases([
    ["0.00", "abc", "abc"],
    ["0.00", "1.5", "1.5"],
    ['0;0;0;"["@"]"', "abc", "[abc]"],
    ['@" units"', "abc", "abc units"],
    ["@@", "ab", "abab"],
    ['0;0;0;"x"', "abc", "x"],
    ['0;"t"@', "abc", "tabc"],
    ["", "abc", "abc"],
    [";;;", "abc", ""],
    ["0;0;0;_(@_)", "abc", " abc "],
  ]);
});

describe("dates", () => {
  cases([
    ["yyyy-mm-dd", 45000, "2023-03-15"],
    ["dddd, mmmm d, yyyy", 45000, "Wednesday, March 15, 2023"],
    ["ddd d mmm yy", 45000, "Wed 15 Mar 23"],
    ["mmmmm", 45000, "M"],
    ["m/d/yyyy", 45000, "3/15/2023"],
    ["dd.mm.yyyy", 45000, "15.03.2023"],
    ["MMMM", 1234, "May"],
    ["dDdD", 1234, "Monday"],
    // 1900-01-00, and Lotus 1-2-3's fictitious 1900-02-29
    ["yyyy-mm-dd", 0, "1900-01-00"],
    ["dddd", 0, "Saturday"],
    ["yyyy-mm-dd", 1, "1900-01-01"],
    ["dddd", 1, "Sunday"],
    ["yyyy-mm-dd", 59, "1900-02-28"],
    ["yyyy-mm-dd", 60, "1900-02-29"],
    ["dddd", 60, "Wednesday"],
    ["yyyy-mm-dd", 61, "1900-03-01"],
    ["yyyy-mm-dd", 2958465, "9999-12-31"],
    ["yyyy-mm-dd", 2958466, ERROR],
    ["yyyy-mm-dd", -1, ERROR],
    ["e", 1, "1900"],
    ["bbbb", 45000, "2566"],
    ["yyyy-mm-dd", 45000.9999999, "2023-03-15"],
    ["[$-F800]dddd, mmmm dd, yyyy", 45000, "Wednesday, March 15, 2023"],
  ]);
});

describe("times", () => {
  cases([
    ["h:mm AM/PM", 0.5, "12:00 PM"],
    ["h:mm AM/PM", 0, "12:00 AM"],
    ["h:mm am/pm", 0.75, "6:00 PM"],
    ["h A/P", 0.25, "6 A"],
    ["h a/p", 0.75, "6 p"],
    ["hh:mm:ss", 0.75, "18:00:00"],
    // Seconds are rounded, minutes are not
    ["h:mm", (10 * 3600 + 30 * 60 + 59) / 86400, "10:30"],
    ["h:mm:ss", (10 * 3600 + 30 * 60 + 59.6) / 86400, "10:31:00"],
    ["h:mm:ss.00", 0.123456, "2:57:46.60"],
    ["hh:mm:ss.000", 0.70707, "16:58:10.848"],
    ["mm:ss", 0.04, "57:36"],
    ["h:m", 0.04, "0:57"],
    // mm alone is the month
    ["mm", 0.04, "01"],
    ["[h]:mm", 1.5, "36:00"],
    ["[mm]:ss", 0.04, "57:36"],
    ["[s]", 0.5, "43200"],
    ["[hhh]", 1, "024"],
    ["[h]:mm:ss", 2.9999999999999996, "72:00:00"],
    ["yyyy-mm-dd hh:mm", 45000.9999999, "2023-03-16 00:00"],
    ["[h]:mm", -0.5, ERROR],
    ["h:mm", -0.5, ERROR],
    ["[$-F400]h:mm", 0.75, "6:00:00 PM"],
  ]);
});

describe("invalid and unsupported formats", () => {
  cases([
    ['"abc', 1, ERROR],
    ["[foo]0", 1, ERROR],
    ["[", 1, ERROR],
    ["0\\", 1, ERROR],
    ["0_", 1, ERROR],
    ["0;0;0;0;0", 1, ERROR],
    ["0;0;0;0", 1, ERROR],
    ["@;@", 1, ERROR],
    ["[>1]0;[>2]0;[>3]0", 1, ERROR],
    ["yyyy 0", 1, ERROR],
    ["h:mm %", 1, ERROR],
    ["0 @", 1, ERROR],
    ["General 0", 1, ERROR],
    ["E+0", 1, ERROR],
    ["0E+", 1, ERROR],
    ["0E0", 1, ERROR],
    ["0/", 1, ERROR],
    ["/0", 1, ERROR],
    ["0.0 ?/?", 1, ERROR],
    ["N0", 1, ERROR],
    ["B", 1, ERROR],
    ["[DBNum1]0", 1, ERROR],
    ["B2yyyy", 1, ERROR],
  ]);
});
