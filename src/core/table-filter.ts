/**
 * Pure helpers for table sorting and filtering (Excel's table AutoFilter):
 * value classification, Excel's sort order, filter matching, validation and
 * cloning of the filter state stored on a TableDefinition.
 */

import type {
  SingleEvaluationResult,
  TableColumnFilter,
  TableColumnFilterState,
  TableFilterCondition,
  TableFilterConditionOperator,
  TableFilterState,
  TableSortDirection,
  TableSortKey,
} from "./types";
import { FormulaError } from "./types";

/** A table cell's value as sorting and filtering see it. */
export type TableCellValue =
  | { kind: "blank" }
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "boolean"; value: boolean }
  | { kind: "error"; error: string };

/** Text compares without regard to case, as Excel's sort and filters do. */
const textCollator = new Intl.Collator("en", {
  sensitivity: "accent",
  numeric: false,
});

export function compareText(left: string, right: string): number {
  return textCollator.compare(left, right);
}

export function tableCellValueFromResult(
  result: SingleEvaluationResult | undefined
): TableCellValue {
  if (!result) {
    return { kind: "blank" };
  }
  if (result.type === "error") {
    return { kind: "error", error: result.err };
  }
  if (result.type === "awaiting-evaluation") {
    return { kind: "error", error: FormulaError.ERROR };
  }
  const value = result.result;
  switch (value.type) {
    case "number":
      return { kind: "number", value: value.value };
    case "infinity":
      return {
        kind: "number",
        value: value.sign === "positive" ? Infinity : -Infinity,
      };
    case "boolean":
      return { kind: "boolean", value: value.value };
    case "string":
      return value.value === ""
        ? { kind: "blank" }
        : { kind: "string", value: value.value };
  }
}

const TYPE_ORDER = { number: 0, string: 1, boolean: 2, error: 3 } as const;

/**
 * Excel's sort order. Ascending: numbers, text (case-insensitive), FALSE before
 * TRUE, errors (all equal). Descending reverses that. Blank cells always sort
 * last, in either direction.
 */
export function compareTableCellValues(
  left: TableCellValue,
  right: TableCellValue,
  direction: TableSortDirection
): number {
  if (left.kind === "blank" || right.kind === "blank") {
    if (left.kind === right.kind) {
      return 0;
    }
    return left.kind === "blank" ? 1 : -1;
  }
  const sign = direction === "descending" ? -1 : 1;
  const typeDifference = TYPE_ORDER[left.kind] - TYPE_ORDER[right.kind];
  if (typeDifference !== 0) {
    return sign * typeDifference;
  }
  switch (left.kind) {
    case "number": {
      const other = (right as { value: number }).value;
      return sign * (left.value === other ? 0 : left.value < other ? -1 : 1);
    }
    case "string":
      return sign * compareText(left.value, (right as { value: string }).value);
    case "boolean":
      return (
        sign * (Number(left.value) - Number((right as { value: boolean }).value))
      );
    case "error":
      return 0;
  }
}

/** A key that groups equal values, as a filter value list does. */
export function tableCellValueKey(value: TableCellValue): string {
  switch (value.kind) {
    case "blank":
      return "blank";
    case "number":
      return `n:${value.value}`;
    case "string":
      // Equal under the collator above: case-insensitive, accent-sensitive.
      return `s:${value.value.toLocaleLowerCase("en")}`;
    case "boolean":
      return `b:${value.value}`;
    case "error":
      return `e:${value.error}`;
  }
}

function equalsFilterValue(
  value: TableCellValue,
  target: string | number | boolean
): boolean {
  switch (typeof target) {
    case "number":
      return value.kind === "number" && value.value === target;
    case "boolean":
      return value.kind === "boolean" && value.value === target;
    case "string":
      return (
        value.kind === "string" && compareText(value.value, target) === 0
      );
  }
}

function lowerText(value: TableCellValue): string | undefined {
  return value.kind === "string" ? value.value.toLocaleLowerCase("en") : undefined;
}

function compareForCondition(
  value: TableCellValue,
  target: string | number | boolean
): number | undefined {
  if (typeof target === "number") {
    return value.kind === "number"
      ? value.value === target
        ? 0
        : value.value < target
        ? -1
        : 1
      : undefined;
  }
  if (typeof target === "string") {
    return value.kind === "string" ? compareText(value.value, target) : undefined;
  }
  return value.kind === "boolean"
    ? Number(value.value) - Number(target)
    : undefined;
}

export function matchesTableFilterCondition(
  value: TableCellValue,
  condition: TableFilterCondition
): boolean {
  const target = condition.value;
  switch (condition.operator) {
    case "equals":
      return target === ""
        ? value.kind === "blank"
        : equalsFilterValue(value, target);
    case "notEquals":
      return !matchesTableFilterCondition(value, {
        operator: "equals",
        value: target,
      });
    case "greaterThan":
    case "greaterThanOrEqual":
    case "lessThan":
    case "lessThanOrEqual": {
      const comparison = compareForCondition(value, target);
      if (comparison === undefined) {
        return false;
      }
      switch (condition.operator) {
        case "greaterThan":
          return comparison > 0;
        case "greaterThanOrEqual":
          return comparison >= 0;
        case "lessThan":
          return comparison < 0;
        case "lessThanOrEqual":
          return comparison <= 0;
      }
      return false;
    }
    case "beginsWith":
    case "endsWith":
    case "contains": {
      const text = lowerText(value);
      if (text === undefined) {
        return false;
      }
      const needle = String(target).toLocaleLowerCase("en");
      return condition.operator === "beginsWith"
        ? text.startsWith(needle)
        : condition.operator === "endsWith"
        ? text.endsWith(needle)
        : text.includes(needle);
    }
    case "notContains":
      return !matchesTableFilterCondition(value, {
        operator: "contains",
        value: target,
      });
  }
}

/** Whether a row with this value in the filtered column stays visible. */
export function matchesTableColumnFilter(
  value: TableCellValue,
  filter: TableColumnFilter
): boolean {
  if (filter.type === "values") {
    switch (value.kind) {
      case "blank":
        return filter.blanks === true;
      case "error":
        return (filter.errors ?? []).includes(value.error);
      default:
        return filter.values.some((target) =>
          equalsFilterValue(value, target)
        );
    }
  }
  if (filter.conditions.length === 0) {
    return true;
  }
  return filter.match === "any"
    ? filter.conditions.some((condition) =>
        matchesTableFilterCondition(value, condition)
      )
    : filter.conditions.every((condition) =>
        matchesTableFilterCondition(value, condition)
      );
}

const CONDITION_OPERATORS = new Set<TableFilterConditionOperator>([
  "equals",
  "notEquals",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
  "beginsWith",
  "endsWith",
  "contains",
  "notContains",
]);

function isFilterValue(value: unknown): value is string | number | boolean {
  return (
    typeof value === "string" ||
    typeof value === "boolean" ||
    (typeof value === "number" && !Number.isNaN(value))
  );
}

/**
 * Validates a column filter and returns a detached copy in canonical form.
 * Throws a descriptive error for anything a filter menu should never send.
 */
export function normalizeTableColumnFilter(
  filter: TableColumnFilter
): TableColumnFilter {
  if (!filter || typeof filter !== "object") {
    throw new Error("A table column filter must be an object");
  }
  if (filter.type === "values") {
    if (!Array.isArray(filter.values) || !filter.values.every(isFilterValue)) {
      throw new Error(
        "A values filter needs `values`: an array of text, numbers and booleans"
      );
    }
    if (filter.errors !== undefined) {
      if (
        !Array.isArray(filter.errors) ||
        !filter.errors.every((error) => typeof error === "string")
      ) {
        throw new Error("A values filter's `errors` must be error codes");
      }
    }
    return {
      type: "values",
      values: [...filter.values],
      ...(filter.blanks ? { blanks: true } : {}),
      ...(filter.errors && filter.errors.length > 0
        ? { errors: [...filter.errors] }
        : {}),
    };
  }
  if (filter.type === "condition") {
    if (!Array.isArray(filter.conditions) || filter.conditions.length === 0) {
      throw new Error("A condition filter needs at least one condition");
    }
    for (const condition of filter.conditions) {
      if (
        !condition ||
        !CONDITION_OPERATORS.has(condition.operator) ||
        !isFilterValue(condition.value)
      ) {
        throw new Error(
          `Invalid filter condition: ${JSON.stringify(condition)}`
        );
      }
    }
    if (
      filter.match !== undefined &&
      filter.match !== "all" &&
      filter.match !== "any"
    ) {
      throw new Error('A condition filter\'s `match` is "all" or "any"');
    }
    return {
      type: "condition",
      conditions: filter.conditions.map((condition) => ({
        operator: condition.operator,
        value: condition.value,
      })),
      ...(filter.match === "any" ? { match: "any" as const } : {}),
    };
  }
  throw new Error(
    `Unknown table column filter type: ${String(
      (filter as { type?: unknown }).type
    )}`
  );
}

/** Validates sort keys against a table's column count. */
export function normalizeTableSortKeys(
  keys: readonly TableSortKey[],
  columnCount: number
): TableSortKey[] {
  if (!Array.isArray(keys) || keys.length === 0) {
    throw new Error("Sorting a table needs at least one sort key");
  }
  const seen = new Set<number>();
  const normalized: TableSortKey[] = [];
  for (const key of keys) {
    assertTableColumn(key?.column, columnCount);
    if (key.direction !== "ascending" && key.direction !== "descending") {
      throw new Error('A sort direction is "ascending" or "descending"');
    }
    // A column used again after its first key cannot change the order.
    if (seen.has(key.column)) {
      continue;
    }
    seen.add(key.column);
    normalized.push({ column: key.column, direction: key.direction });
  }
  return normalized;
}

export function assertTableColumn(column: unknown, columnCount: number): void {
  if (
    typeof column !== "number" ||
    !Number.isInteger(column) ||
    column < 0 ||
    column >= columnCount
  ) {
    throw new Error(
      `Table column ${String(column)} is out of range: the table has ${columnCount} column${
        columnCount === 1 ? "" : "s"
      }`
    );
  }
}

export function cloneTableColumnFilter(
  filter: TableColumnFilter
): TableColumnFilter {
  return filter.type === "values"
    ? {
        type: "values",
        values: [...filter.values],
        ...(filter.blanks !== undefined ? { blanks: filter.blanks } : {}),
        ...(filter.errors !== undefined ? { errors: [...filter.errors] } : {}),
      }
    : {
        type: "condition",
        conditions: filter.conditions.map((condition) => ({ ...condition })),
        ...(filter.match !== undefined ? { match: filter.match } : {}),
      };
}

export function cloneTableFilterState(
  state: TableFilterState
): TableFilterState {
  return {
    columns: state.columns.map((column) => ({
      column: column.column,
      filter: cloneTableColumnFilter(column.filter),
    })),
    hiddenRows: [...state.hiddenRows],
  };
}

export function cloneTableSortKeys(keys: readonly TableSortKey[]): TableSortKey[] {
  return keys.map((key) => ({ column: key.column, direction: key.direction }));
}

function columnFiltersEqual(
  left: TableColumnFilter,
  right: TableColumnFilter
): boolean {
  if (left.type === "values" && right.type === "values") {
    return (
      left.values.length === right.values.length &&
      left.values.every((value, index) =>
        Object.is(value, right.values[index])
      ) &&
      (left.blanks ?? false) === (right.blanks ?? false) &&
      (left.errors ?? []).length === (right.errors ?? []).length &&
      (left.errors ?? []).every((error, index) => error === right.errors![index])
    );
  }
  if (left.type === "condition" && right.type === "condition") {
    return (
      (left.match ?? "all") === (right.match ?? "all") &&
      left.conditions.length === right.conditions.length &&
      left.conditions.every(
        (condition, index) =>
          condition.operator === right.conditions[index]!.operator &&
          Object.is(condition.value, right.conditions[index]!.value)
      )
    );
  }
  return false;
}

export function tableFilterStatesEqual(
  left: TableFilterState | undefined,
  right: TableFilterState | undefined
): boolean {
  if (!left || !right) {
    return left === right;
  }
  return (
    left.columns.length === right.columns.length &&
    left.columns.every(
      (column, index) =>
        column.column === right.columns[index]!.column &&
        columnFiltersEqual(column.filter, right.columns[index]!.filter)
    ) &&
    left.hiddenRows.length === right.hiddenRows.length &&
    left.hiddenRows.every((row, index) => row === right.hiddenRows[index])
  );
}

export function tableSortKeysEqual(
  left: readonly TableSortKey[] | undefined,
  right: readonly TableSortKey[] | undefined
): boolean {
  if (!left || !right) {
    return left === right;
  }
  return (
    left.length === right.length &&
    left.every(
      (key, index) =>
        key.column === right[index]!.column &&
        key.direction === right[index]!.direction
    )
  );
}

/** Orders filtered columns by column index, one entry per column. */
export function sortTableColumnFilters(
  columns: readonly TableColumnFilterState[]
): TableColumnFilterState[] {
  return [...columns].sort((left, right) => left.column - right.column);
}
