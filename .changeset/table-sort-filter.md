---
"@ricsam/formula-engine": minor
---

Sort and filter tables like Excel tables. `sortTable` reorders a table's data
rows by one or more columns, `setTableColumnFilter` hides the rows a column's
criteria leave out, and `getHiddenRows` tells a grid which rows to draw with
zero height. Neither changes what a cell address means: sorting moves data,
filtering only hides rows.

```ts
engine.sortTable({
  workbookName: "Budget",
  tableName: "Orders",
  keys: [{ column: 2, direction: "descending" }],
});
engine.setTableColumnFilter({
  workbookName: "Budget",
  tableName: "Orders",
  column: 0,
  filter: { type: "values", values: ["North", "South"] },
});
engine.getHiddenRows({ workbookName: "Budget", sheetName: "Q1" }); // Set {3, 5}
```

A sort is stable and uses Excel's order: numbers, text without regard to case,
`FALSE`, `TRUE`, errors, with blank cells last in either direction. Each row's
content, cell metadata, direct cell styles and data types move together, and
moved formulas are adjusted like pasted ones, so `=B5*C5` moved to row 2 reads
`=B2*C2` while `A:A` keeps spanning its column. Formulas elsewhere that point
into the table are not rewritten, as in Excel. A sort is one undo step.

A column filter is a value list (`{ type: "values", values, blanks?, errors? }`,
text matched without regard to case) or conditions (`equals`, `notEquals`,
`greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`,
`beginsWith`, `endsWith`, `contains`, `notContains`, combined with `all` or
`any`). Filters apply when they change, after a sort and when `updateTable`
resizes the table, not on every edit; `reapplyTableFilters` applies them again
and `clearTableFilters` removes them. `getTableColumnValues` lists a column's
distinct values with counts for a filter menu, narrowed by the other columns'
filters like Excel's list.

`TableDefinition` gains `filter` (the criteria and the rows they hid) and `sort`
(the last sort's keys). Both are undoable, serialized with the engine, copied
with a cloned sheet and moved with a moved table; a resize keeps a column's
filter by header name. `WorkbookData` tables accept `filters` and `sort`.
