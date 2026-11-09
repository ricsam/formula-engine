import { functionDefinitions } from "../functions/function-registry";

export type FormulaFunctionCategory =
  | "array"
  | "information"
  | "logical"
  | "lookup"
  | "math"
  | "reference"
  | "text";

export interface FormulaFunctionParameter {
  /** The parameter label used in signatures and editor placeholders. */
  readonly name: string;
  /** Whether callers may omit this parameter. */
  readonly optional: boolean;
  /** Whether this parameter, or the group it belongs to, may repeat. */
  readonly repeatable: boolean;
}

export interface FormulaFunctionDescriptor {
  /** The canonical, case-insensitive function name. */
  readonly name: string;
  /** Additional callable names for the same function. */
  readonly aliases: readonly string[];
  readonly category: FormulaFunctionCategory;
  /** A short, plain-text summary suitable for completion details. */
  readonly description: string;
  /** A display signature using square brackets for optional arguments. */
  readonly signature: string;
  readonly parameters: readonly FormulaFunctionParameter[];
}

type FormulaFunctionMetadata = Omit<
  FormulaFunctionDescriptor,
  "name" | "aliases"
>;

const parameter = (
  name: string,
  optional = false,
  repeatable = false
): FormulaFunctionParameter => ({ name, optional, repeatable });

const functionMetadata = {
  ADDRESS: {
    category: "reference",
    description: "Creates a cell reference as text from row and column numbers.",
    signature: "ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])",
    parameters: [
      parameter("row_num"),
      parameter("column_num"),
      parameter("abs_num", true),
      parameter("a1", true),
      parameter("sheet_text", true),
    ],
  },
  AND: {
    category: "logical",
    description: "Returns TRUE when every argument is truthy.",
    signature: "AND(logical1, [logical2], ...)",
    parameters: [
      parameter("logical1"),
      parameter("logical2", true, true),
    ],
  },
  AVERAGE: {
    category: "math",
    description: "Returns the arithmetic mean of numeric values.",
    signature: "AVERAGE(value1, [value2], ...)",
    parameters: [parameter("value1"), parameter("value2", true, true)],
  },
  AVERAGEIF: {
    category: "math",
    description: "Averages cells whose corresponding values meet a criterion.",
    signature: "AVERAGEIF(range, criteria, [average_range])",
    parameters: [
      parameter("range"),
      parameter("criteria"),
      parameter("average_range", true),
    ],
  },
  AVERAGEIFS: {
    category: "math",
    description: "Averages cells that meet every supplied criterion.",
    signature:
      "AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    parameters: [
      parameter("average_range"),
      parameter("criteria_range1"),
      parameter("criteria1"),
      parameter("criteria_range2", true, true),
      parameter("criteria2", true, true),
    ],
  },
  CEILING: {
    category: "math",
    description: "Rounds a number up to the nearest multiple of significance.",
    signature: "CEILING(number, significance)",
    parameters: [parameter("number"), parameter("significance")],
  },
  CELL: {
    category: "information",
    description: "Returns requested information about a cell.",
    signature: "CELL(info_type, [reference])",
    parameters: [parameter("info_type"), parameter("reference", true)],
  },
  COLUMN: {
    category: "information",
    description: "Returns the column number of a reference or the formula cell.",
    signature: "COLUMN([reference])",
    parameters: [parameter("reference", true)],
  },
  COLUMNS: {
    category: "information",
    description: "Returns the number of columns in a reference or array.",
    signature: "COLUMNS(array)",
    parameters: [parameter("array")],
  },
  CONCATENATE: {
    category: "text",
    description: "Joins values into one text string.",
    signature: "CONCATENATE(text1, [text2], ...)",
    parameters: [parameter("text1"), parameter("text2", true, true)],
  },
  COUNT: {
    category: "lookup",
    description: "Counts numeric values in the supplied arguments.",
    signature: "COUNT(value1, [value2], ...)",
    parameters: [parameter("value1"), parameter("value2", true, true)],
  },
  COUNTA: {
    category: "lookup",
    description: "Counts values that are not empty, including text, logical values and errors.",
    signature: "COUNTA(value1, [value2], ...)",
    parameters: [parameter("value1"), parameter("value2", true, true)],
  },
  COUNTIF: {
    category: "lookup",
    description: "Counts cells in a range that meet a criterion.",
    signature: "COUNTIF(range, criteria)",
    parameters: [parameter("range"), parameter("criteria")],
  },
  COUNTIFS: {
    category: "lookup",
    description: "Counts cells that meet every supplied criterion.",
    signature:
      "COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    parameters: [
      parameter("criteria_range1"),
      parameter("criteria1"),
      parameter("criteria_range2", true, true),
      parameter("criteria2", true, true),
    ],
  },
  EXACT: {
    category: "text",
    description: "Tests whether two text values are exactly equal.",
    signature: "EXACT(text1, text2)",
    parameters: [parameter("text1"), parameter("text2")],
  },
  FILTER: {
    category: "array",
    description:
      "Returns the rows or columns of an array for which a condition is true.",
    signature: "FILTER(array, include, [if_empty])",
    parameters: [
      parameter("array"),
      parameter("include"),
      parameter("if_empty", true),
    ],
  },
  FIND: {
    category: "text",
    description: "Returns the position of one text value within another.",
    signature: "FIND(find_text, within_text, [start_num])",
    parameters: [
      parameter("find_text"),
      parameter("within_text"),
      parameter("start_num", true),
    ],
  },
  HSTACK: {
    category: "array",
    description: "Joins arrays side by side into one array.",
    signature: "HSTACK(array1, [array2], ...)",
    parameters: [parameter("array1"), parameter("array2", true, true)],
  },
  IF: {
    category: "logical",
    description: "Returns one value for a true condition and another for false.",
    signature: "IF(logical_test, value_if_true, [value_if_false])",
    parameters: [
      parameter("logical_test"),
      parameter("value_if_true"),
      parameter("value_if_false", true),
    ],
  },
  IFERROR: {
    category: "logical",
    description: "Returns a fallback value when evaluating a value produces an error.",
    signature: "IFERROR(value, value_if_error)",
    parameters: [parameter("value"), parameter("value_if_error")],
  },
  IFS: {
    category: "logical",
    description: "Returns the value that belongs to the first true condition.",
    signature:
      "IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)",
    parameters: [
      parameter("logical_test1"),
      parameter("value_if_true1"),
      parameter("logical_test2", true, true),
      parameter("value_if_true2", true, true),
    ],
  },
  INDEX: {
    category: "lookup",
    description: "Returns a value at a row and column within a range or array.",
    signature: "INDEX(array, row_num, [column_num])",
    parameters: [
      parameter("array"),
      parameter("row_num"),
      parameter("column_num", true),
    ],
  },
  INDIRECT: {
    category: "reference",
    description: "Resolves a reference supplied as text.",
    signature: "INDIRECT(ref_text, [a1])",
    parameters: [parameter("ref_text"), parameter("a1", true)],
  },
  INT: {
    category: "math",
    description: "Rounds a number down to the nearest integer.",
    signature: "INT(number)",
    parameters: [parameter("number")],
  },
  ISBLANK: {
    category: "information",
    description: "Returns TRUE when a value is blank.",
    signature: "ISBLANK(value)",
    parameters: [parameter("value")],
  },
  ISNUMBER: {
    category: "information",
    description: "Returns TRUE when a value is a number.",
    signature: "ISNUMBER(value)",
    parameters: [parameter("value")],
  },
  LEFT: {
    category: "text",
    description: "Returns characters from the beginning of a text value.",
    signature: "LEFT(text, [num_chars])",
    parameters: [parameter("text"), parameter("num_chars", true)],
  },
  LEN: {
    category: "text",
    description: "Returns the number of characters in a text value.",
    signature: "LEN(text)",
    parameters: [parameter("text")],
  },
  LET: {
    category: "logical",
    description:
      "Names intermediate results and returns a final calculation that can use those names.",
    signature:
      "LET(name1, name_value1, calculation_or_name2, [name_value2, calculation_or_name3], ...)",
    parameters: [
      parameter("name1"),
      parameter("name_value1"),
      parameter("calculation_or_name2"),
      parameter("name_value2", true, true),
      parameter("calculation_or_name3", true, true),
    ],
  },
  MATCH: {
    category: "lookup",
    description: "Returns the relative position of a value in a range or array.",
    signature: "MATCH(lookup_value, lookup_array, [match_type])",
    parameters: [
      parameter("lookup_value"),
      parameter("lookup_array"),
      parameter("match_type", true),
    ],
  },
  MAX: {
    category: "math",
    description: "Returns the largest numeric value.",
    signature: "MAX(value1, [value2], ...)",
    parameters: [parameter("value1"), parameter("value2", true, true)],
  },
  MAXIF: {
    category: "math",
    description: "Returns the largest value whose corresponding value meets a criterion.",
    signature: "MAXIF(range, criteria, [max_range])",
    parameters: [
      parameter("range"),
      parameter("criteria"),
      parameter("max_range", true),
    ],
  },
  MAXIFS: {
    category: "math",
    description: "Returns the largest value that meets every supplied criterion.",
    signature:
      "MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    parameters: [
      parameter("max_range"),
      parameter("criteria_range1"),
      parameter("criteria1"),
      parameter("criteria_range2", true, true),
      parameter("criteria2", true, true),
    ],
  },
  MID: {
    category: "text",
    description: "Returns a requested number of characters from within text.",
    signature: "MID(text, start_num, num_chars)",
    parameters: [
      parameter("text"),
      parameter("start_num"),
      parameter("num_chars"),
    ],
  },
  MIN: {
    category: "math",
    description: "Returns the smallest numeric value.",
    signature: "MIN(value1, [value2], ...)",
    parameters: [parameter("value1"), parameter("value2", true, true)],
  },
  MINIF: {
    category: "math",
    description: "Returns the smallest value whose corresponding value meets a criterion.",
    signature: "MINIF(range, criteria, [min_range])",
    parameters: [
      parameter("range"),
      parameter("criteria"),
      parameter("min_range", true),
    ],
  },
  MINIFS: {
    category: "math",
    description: "Returns the smallest value that meets every supplied criterion.",
    signature:
      "MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    parameters: [
      parameter("min_range"),
      parameter("criteria_range1"),
      parameter("criteria1"),
      parameter("criteria_range2", true, true),
      parameter("criteria2", true, true),
    ],
  },
  OFFSET: {
    category: "reference",
    description: "Returns a reference offset from a starting reference.",
    signature: "OFFSET(reference, rows, cols, [height], [width])",
    parameters: [
      parameter("reference"),
      parameter("rows"),
      parameter("cols"),
      parameter("height", true),
      parameter("width", true),
    ],
  },
  OR: {
    category: "logical",
    description: "Returns TRUE when any argument is truthy.",
    signature: "OR(logical1, [logical2], ...)",
    parameters: [
      parameter("logical1"),
      parameter("logical2", true, true),
    ],
  },
  RIGHT: {
    category: "text",
    description: "Returns characters from the end of a text value.",
    signature: "RIGHT(text, [num_chars])",
    parameters: [parameter("text"), parameter("num_chars", true)],
  },
  ROUND: {
    category: "math",
    description:
      "Rounds a number to a number of decimal places, halves away from zero.",
    signature: "ROUND(number, num_digits)",
    parameters: [parameter("number"), parameter("num_digits")],
  },
  ROUNDDOWN: {
    category: "math",
    description: "Rounds a number toward zero to a number of decimal places.",
    signature: "ROUNDDOWN(number, num_digits)",
    parameters: [parameter("number"), parameter("num_digits")],
  },
  ROUNDUP: {
    category: "math",
    description: "Rounds a number away from zero to a number of decimal places.",
    signature: "ROUNDUP(number, num_digits)",
    parameters: [parameter("number"), parameter("num_digits")],
  },
  ROW: {
    category: "information",
    description: "Returns the row number of a reference or the formula cell.",
    signature: "ROW([reference])",
    parameters: [parameter("reference", true)],
  },
  ROWS: {
    category: "information",
    description: "Returns the number of rows in a reference or array.",
    signature: "ROWS(array)",
    parameters: [parameter("array")],
  },
  SEQUENCE: {
    category: "array",
    description: "Generates a spilled array of sequential numbers.",
    signature: "SEQUENCE(rows, [columns], [start], [step])",
    parameters: [
      parameter("rows"),
      parameter("columns", true),
      parameter("start", true),
      parameter("step", true),
    ],
  },
  "STDEV.S": {
    category: "math",
    description: "Returns the standard deviation of a sample of numbers.",
    signature: "STDEV.S(number1, [number2], ...)",
    parameters: [parameter("number1"), parameter("number2", true, true)],
  },
  SUM: {
    category: "math",
    description: "Adds numeric values.",
    signature: "SUM(value1, [value2], ...)",
    parameters: [parameter("value1"), parameter("value2", true, true)],
  },
  SUMIF: {
    category: "math",
    description: "Adds cells whose corresponding values meet a criterion.",
    signature: "SUMIF(range, criteria, [sum_range])",
    parameters: [
      parameter("range"),
      parameter("criteria"),
      parameter("sum_range", true),
    ],
  },
  SUMIFS: {
    category: "math",
    description: "Adds cells that meet every supplied criterion.",
    signature:
      "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
    parameters: [
      parameter("sum_range"),
      parameter("criteria_range1"),
      parameter("criteria1"),
      parameter("criteria_range2", true, true),
      parameter("criteria2", true, true),
    ],
  },
  SUMPRODUCT: {
    category: "math",
    description:
      "Multiplies corresponding cells of same-sized arrays and adds the products.",
    signature: "SUMPRODUCT(array1, [array2], ...)",
    parameters: [parameter("array1"), parameter("array2", true, true)],
  },
  TEXT: {
    category: "text",
    description:
      "Formats a number, date or time as text using a number format code.",
    signature: "TEXT(value, format_text)",
    parameters: [parameter("value"), parameter("format_text")],
  },
  TEXTJOIN: {
    category: "text",
    description: "Joins text values using a delimiter, optionally ignoring empty values.",
    signature: "TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)",
    parameters: [
      parameter("delimiter"),
      parameter("ignore_empty"),
      parameter("text1"),
      parameter("text2", true, true),
    ],
  },
  TOCOL: {
    category: "array",
    description:
      "Returns an array as a single column, optionally skipping blanks and errors.",
    signature: "TOCOL(array, [ignore], [scan_by_column])",
    parameters: [
      parameter("array"),
      parameter("ignore", true),
      parameter("scan_by_column", true),
    ],
  },
  TOROW: {
    category: "array",
    description:
      "Returns an array as a single row, optionally skipping blanks and errors.",
    signature: "TOROW(array, [ignore], [scan_by_column])",
    parameters: [
      parameter("array"),
      parameter("ignore", true),
      parameter("scan_by_column", true),
    ],
  },
  TRANSPOSE: {
    category: "array",
    description: "Swaps the rows and columns of a range or array.",
    signature: "TRANSPOSE(array)",
    parameters: [parameter("array")],
  },
  UNIQUE: {
    category: "array",
    description:
      "Returns the distinct rows, or columns, of a range or array.",
    signature: "UNIQUE(array, [by_col], [exactly_once])",
    parameters: [
      parameter("array"),
      parameter("by_col", true),
      parameter("exactly_once", true),
    ],
  },
  VLOOKUP: {
    category: "lookup",
    description:
      "Finds a value in the first column of a table and returns a value from the same row.",
    signature: "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
    parameters: [
      parameter("lookup_value"),
      parameter("table_array"),
      parameter("col_index_num"),
      parameter("range_lookup", true),
    ],
  },
  XLOOKUP: {
    category: "lookup",
    description: "Finds a value and returns the corresponding value from another array.",
    signature:
      "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
    parameters: [
      parameter("lookup_value"),
      parameter("lookup_array"),
      parameter("return_array"),
      parameter("if_not_found", true),
      parameter("match_mode", true),
      parameter("search_mode", true),
    ],
  },
} satisfies Record<keyof typeof functionDefinitions, FormulaFunctionMetadata>;

type RegisteredFunctionName = keyof typeof functionDefinitions;

const formulaFunctionCatalog: readonly FormulaFunctionDescriptor[] =
  Object.freeze(
    (Object.keys(functionDefinitions) as RegisteredFunctionName[])
      .sort((left, right) => left.localeCompare(right))
      .map((registeredName) => {
        const definition = functionDefinitions[registeredName];
        const metadata = functionMetadata[registeredName];

        return Object.freeze({
          name: definition.name,
          aliases: Object.freeze([...(definition.aliases ?? [])]),
          category: metadata.category,
          description: metadata.description,
          signature: metadata.signature,
          parameters: Object.freeze(
            metadata.parameters.map((item) => Object.freeze({ ...item }))
          ),
        });
      })
  );

const formulaFunctionsByName = new Map<string, FormulaFunctionDescriptor>();

for (const descriptor of formulaFunctionCatalog) {
  formulaFunctionsByName.set(descriptor.name.toUpperCase(), descriptor);
  for (const alias of descriptor.aliases) {
    formulaFunctionsByName.set(alias.toUpperCase(), descriptor);
  }
}

/**
 * Returns metadata for every canonical function registered by FormulaEngine.
 * The returned catalog and its entries are immutable and sorted by name.
 */
export function getFormulaFunctionCatalog(): readonly FormulaFunctionDescriptor[] {
  return formulaFunctionCatalog;
}

/** Finds canonical function metadata by name or alias, case-insensitively. */
export function findFormulaFunction(
  name: string
): FormulaFunctionDescriptor | undefined {
  return formulaFunctionsByName.get(name.toUpperCase());
}
