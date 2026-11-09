# @ricsam/formula-engine

## 0.5.0

### Minor Changes

- 6993906: Reference tables in other workbooks as `[Workbook]Table[Column]`, the way
  `[Workbook]Sheet!A1` references a cell there:
  `=INDEX([Batch 1]Results[Yield], 2)` or
  `=SUM([Batch 1]Results[Yield])+SUM([Batch 2]Results[Yield])`. Every structured
  reference form works: `[Workbook]Table[@Column]` (the row of the formula cell),
  column spans such as `[Workbook]Table[[Sample]:[Yield]]`, and the `#Headers`,
  `#Data` and `#All` selectors with or without columns. These formulas were
  syntax errors that showed `#ERROR!`. As before, a reference without the
  workbook names a table in the formula's own workbook.

  The references are tracked like cross-workbook cell references. They
  recalculate when the table's data or size changes, a reference such as
  `INDEX([Batch 1]Results[Yield], 1)` is `#REF!` while the workbook, table or
  column is missing and resolves again once it is added, and renaming the
  workbook, the table or a column header rewrites them. Formula analysis
  resolves them to the other workbook's table and marks all of their tokens
  `external`, like other references into another workbook.

  Breaking: the previous spelling `[Workbook]!Table[Column]` is now a syntax
  error whose message names the new spelling, and formulas are stored as
  `[Workbook]Table[Column]`. Remove the `!` from stored formulas that use it. A
  warm snapshot saved with such a reference restores cold, with a
  `[FormulaEngine snapshot]` warning.

- 50a205e: Add the functions the Operations workbooks evaluated to `#NAME?`: `IFS`,
  `FILTER`, `TRANSPOSE`, `STDEV.S` (alias `STDEV`), `ROUND`, `ROUNDUP`,
  `ROUNDDOWN`, `ISNUMBER`, `ISBLANK`, `SUMPRODUCT`, `VLOOKUP` and `TEXT`, for 59
  functions in all. `GETPIVOTDATA` is not added: the engine has no pivot tables.

  `IFS(test1, value1, ...)` returns the value of the first true condition and
  `#N/A` when none is. Conditions are tested like `IF`'s and only the chosen
  value is evaluated; an array condition chooses cell by cell and broadcasts like
  `IF`. An odd number of arguments is `#VALUE!`.

  `FILTER(array, include, [if_empty])` returns the rows of `array` selected by a
  column `include`, or its columns for a row `include`. Nothing selected is
  `if_empty`, or `#CALC!` without it; an error in `include` is the result. Whole
  columns such as `FILTER(B:B, A:A="x")` work. Booleans are not numbers in this
  engine, so Excel's `(A2:A9="x")*(B2:B9>5)` is `#VALUE!`; combine conditions
  with `IF(A2:A9="x", B2:B9>5)`. `TRANSPOSE(array)` swaps rows and columns lazily
  and spills.

  `ROUND`, `ROUNDUP` and `ROUNDDOWN` round the decimal value Excel shows, with 15
  significant digits, so `ROUND(2.675, 2)` is 2.68, `ROUNDUP(0.1+0.2, 1)` is 0.3
  and `ROUNDDOWN(4.35*100, 0)` is 435. A negative `num_digits` rounds to tens,
  hundreds and so on. Like `INT`, they accept only numbers, and array arguments
  broadcast.

  `STDEV.S` is the sample standard deviation of the numbers in its arguments,
  ranges and spills, such as `STDEV.S(FILTER(...))`. Errors propagate and fewer
  than two numbers is `#DIV/0!`. `SUMPRODUCT` adds the products of same-sized
  arrays: a different size is `#VALUE!`, a cell that is not a number counts as 0,
  and an error anywhere is the result. The `--(...)` idiom is `#VALUE!` here;
  choose values with `IF` instead.

  `VLOOKUP` searches the first column of a table like `MATCH`: exact for `FALSE`
  or `0` (text ignores case, no wildcards) and, for `TRUE`, the default, with
  `MATCH`'s binary search, which finds the largest value less than or equal to
  the lookup value in a sorted column and Excel's row in an unsorted one.
  `col_index_num` below 1 is `#VALUE!` and beyond the table `#REF!`, and a blank
  result cell is `""`. Array lookup values and column numbers spill.

  `ISNUMBER` and `ISBLANK` test a value without propagating errors and test
  arrays element by element. Aggregates read every cell of a bounded result, so
  `SUM(IF(ISBLANK(A1:A9), 1, 0))` counts the empty cells. `ISBLANK` is `TRUE` for
  everything that reads as `""`, including a formula returning `""` and `""`
  itself, where Excel is `TRUE` only for an empty cell.

  `TEXT(value, format_text)` formats numbers, dates and times with Excel number
  format codes: digit placeholders, thousands separators and scaling,
  percentages, scientific notation, fractions, `General`, literal text, up to
  four sections with conditions, currency tags, and date and time codes on
  Excel's 1900 date system, including elapsed time. Numbers are rounded as Excel
  shows them (`TEXT(2.675,"0.00")` is "2.68"). Following the engine's strict
  typing, text is never converted to a number, an empty cell formats as `""`, and
  `format_text` must be text. Array arguments broadcast.

### Patch Changes

- 6993906: Recalculate references to a workbook's tables when `cloneWorkbook` creates it.
  A reference written before the clone, such as `=INDEX([Copy]Results[Yield], 1)`,
  stayed `#REF!`, and formulas in the clone that refer to its own tables could
  reuse that stale result.
- 898e688: Fix unqualified current-row column spans such as `[@[pep_concentration]:[gate]]`,
  which read only their first column. Formatting wrote the span as
  `[@pep_concentration]`, and since the engine normalizes formula content when it
  stores it, `=TEXTJOIN(", ", TRUE, [@[pep_concentration]:[gate]])` returned
  `"10uM"` instead of `"10uM, RFP+"` and the stored formula lost the rest of the
  span. `SUM([@[a]:[b]])` and other functions over such spans were affected the
  same way. Table-qualified spans such as `Table1[@[a]:[b]]` already worked.
- 6993906: Parse workbook names of any length. A workbook qualifier whose name had more
  than eight words, or a whole-sheet alias such as `[Book]Sheet1` followed by
  many more tokens, as in `SUM([Book]Sheet1, 1, 2, 3, 4, 5)`, was a syntax error.
- b597c5a: Search like Excel in `MATCH` with match type `1` (the default) and `-1`. Excel
  does a binary search there without checking the order, and so does `MATCH`
  now, so unsorted data returns the position Excel returns instead of `#N/A`.
  Workbooks that leave out the `,0` of an exact match, such as
  `=INDEX(MAPcondition[MAP Payload],MATCH([@MAP ID],MAPcondition[MAP ID]),0)`
  over unsorted IDs, now give Excel's results: with `MAP-994`, `MAP-10977` and
  `MAP-320` in `A1:A3`, `=MATCH("MAP-994",A1:A3)` is `3`, as in Excel. Sorted
  data finds the same positions as before. Blank cells, errors and values of
  another type are stepped over the way Excel steps over them, so
  `=MATCH(9.99E+307,B:B)` finds the last number in column B whatever the order.
  An unbounded range such as `B:B` is searched up to its last non-empty cell.
- 6993906: Scope table and sheet renames to their workbook. Renaming a table or sheet
  rewrote every reference to that name in every workbook, so with two batch
  workbooks that each have a `Results` table or a `Data` sheet, renaming one
  broke `=SUM(Results[Yield])` and `=Data!A1` in the other, and even references
  qualified with the other workbook's name. Only references that resolve to the
  renamed workbook's table or sheet change now. An unqualified reference in a
  global named expression, which can resolve in any workbook, is still renamed.
- 6993906: A formula that does not parse shows `#ERROR!`. Some syntax errors showed
  another error value because their message mentioned it, such as `#REF!` for
  `=Sheet1!` ("Expected cell reference after Sheet1!").
- 025424f: Keep every table column when `setSheetContent` reorders a table's header names.
  The columns were renamed one at a time, so writing `Price, Qty` over the header
  row `Qty, Price` gave both columns the name `Price` midway and the table lost
  one of them; structured references to the lost column read as empty. The
  columns are now renamed together, and references follow their columns as with
  any other header edit. This affected writing a dataset whose columns come in a
  different order into a table, for example with `updateTable` and
  `setSheetContent` in one `transact` call.

## 0.4.0

### Minor Changes

- Add `LET`, `TOCOL`, `TOROW`, `HSTACK`, `ROWS`, `COLUMNS` (alias `COLS`), `INT`
  and `COUNTA`, for 47 functions in all. Together with omitted arguments and
  broadcasting they support combinatorial spills such as every pair of distinct
  values from two columns:
  `=LET(a, UNIQUE(TOCOL(A2:A1000, 1)), b, UNIQUE(TOCOL(B2:B1000, 1)), HSTACK(TOCOL(IF(SEQUENCE(,ROWS(b)), a)), TOCOL(IF(SEQUENCE(ROWS(a)), TOROW(b)))))`.

  `LET(name1, name_value1, ..., calculation)` names intermediate results. Names
  are local to the formula and case-insensitive, each value sees the names
  declared before it, a nested `LET` can declare an outer name again, and a
  `LET` name hides a named expression of the same name. A name bound to a
  reference stays a reference. A wrong argument count or a missing calculation
  is `#VALUE!`, an invalid or repeated name is `#NAME?`, and a `LET` whose
  expansion exceeds 10,000 nodes is `#CALC!`. Unlike Excel, a name used n times
  evaluates its value n times per calculation, an unused value is never
  evaluated, and an omitted value is 0.

  `TOCOL(array, [ignore], [scan_by_column])` and `TOROW` flatten a range or
  array into one column or row, optionally skipping blanks (1), errors (2) or
  both (3). Unbounded ranges such as `A:A` work; skipping blanks drops their
  empty tail. Nothing left is `#CALC!`. `HSTACK(array1, ...)` appends arrays
  side by side and pads shorter ones with `#N/A`.

  `ROWS` and `COLUMNS` measure a range without reading its cells and a spilled
  result by its spill area. A reference to a spilling cell measures the whole
  spill, and an unbounded dimension is `INFINITY`. `INT` rounds down toward
  negative infinity, element by element over arrays, and accepts only numbers.
  `COUNTA` counts non-empty values, including errors and, inside a range,
  formulas that return `""`.

- Broadcast arrays like Excel in operators and `IF`. In each dimension the
  result is as large as the largest operand; an operand of size 1 is repeated
  and cells beyond a shorter operand are `#N/A`. `SEQUENCE(3)*SEQUENCE(,4)`
  spills a 3 × 4 table and `IF(SEQUENCE(,3), A1:A2)` spills 2 × 3. An array
  condition makes `IF` choose per cell and read only the chosen branch, and
  `IFERROR` checks an array cell by cell, broadcasting an array fallback. Array
  literals can be read whole. Aggregates now accept these results:
  `SUM(IF(...))`, `COUNT(IF(...))`, `UNIQUE(IF(...))`, `SUM(IFERROR(...))` and
  `SUM({1,2,3})` work instead of returning `#ERROR!`.

  Behavior changes:

  - Two arrays combine by broadcasting instead of by position. Cells beyond a
    shorter array are `#N/A` without being read; before, `A1:A3+B1:B2` read
    `B3`, and mismatched shapes gave errors such as `#REF!`.
  - A blank cell inside an array behaves as it does in a single-cell formula:
    `&` treats it as empty text and arithmetic returns `#VALUE!`. Before, a
    blank cell in an operation between two arrays gave `#N/A`.
  - When an aggregate reads an operator result, a position filled in only one
    operand is computed with a blank instead of being `#N/A`, and `#N/A` cells
    beyond a shorter array are included. A position where every array operand
    is blank is still left out.
  - An error in a single-value `IF` branch or `IFERROR` fallback appears only in
    the cells that use it instead of failing the whole result.
  - An error in one element of an array literal is that cell's value instead of
    the whole array's, and `{}` is `#CALC!`.

- Accept omitted function arguments such as `SEQUENCE(,3)`, `SUM(1,,2)` and
  `IF(A1,1,)`. Formulas like these were syntax errors that showed `#ERROR!`;
  they now evaluate and are stored in compact form like other valid formulas.
  Array literals still reject omitted elements.

  An omitted argument evaluates like a reference to an empty cell, except where
  a function follows Excel's rule for it. `SEQUENCE` uses its defaults for rows,
  columns, start and step. `IF`, `IFERROR` and `CEILING` read an omitted value
  as 0. `UNIQUE`, `XLOOKUP`, `OFFSET`'s height and width and `ADDRESS` use
  their defaults. `MATCH(x, r,)` is an exact match, `OFFSET(A1,,1)` moves 0
  rows, `LEFT("abc",)`, `RIGHT` and `MID` take 0 characters, `FIND` and `MID`
  with an omitted start position return `#VALUE!`, and `TEXTJOIN` reads an
  omitted delimiter as empty text and an omitted `ignore_empty` as `TRUE`.
  Elsewhere, results can differ from Excel, which reads most omitted arguments
  as 0: `COUNT(1,,2)` is 2 here and 3 in Excel.

  Related function changes:

  - `INDEX` accepts a row or column number of 0, and an omitted one counts as
    0: `INDEX(A1:C3,,2)` is the whole second column. A selection from a range
    reads like that range; from a computed array it spills. Before,
    `INDEX(A1:A2,0,1)` was `#VALUE!`.
  - `ADDRESS` produces R1C1 text when `a1` is `FALSE`: `ADDRESS(2,3,4,FALSE)` is
    `"R[2]C[3]"`. Before, it returned `#VALUE!`.
  - `OFFSET` returns `#VALUE!` for a non-numeric `height` or `width`, and
    `ADDRESS` for a non-numeric `abs_num`, instead of ignoring them.
  - `SEQUENCE` returns `#VALUE!` for an array argument and for `-INFINITY`
    instead of `#ERROR!` and `1`, and `SEQUENCE(INFINITY, 2)` fills row by row
    instead of repeating each number across the row.

  `formatFormula` prints an omitted argument as nothing, as in `=SEQUENCE(, 3)`
  and `=IF(A1, 1,)`. Pretty output gives an omitted last argument no line of its
  own and keeps `{}` on one line.

- Add `LET` support to the formula tooling. Formula analysis reports `LET` names
  as the new public `FormulaTokenKind` `"variable"`, at the declaration and
  every use in scope, without a reference entry or a missing-name warning, and
  scopes the names of a `LET` that is still being typed. The new warnings
  `let.invalid-name` and `let.duplicate-name` flag names that `LET` rejects.
  Editor integrations that switch over `FormulaTokenKind` need to handle the
  new kind.

  Renaming a named expression is scope-aware. `LET` names spelled like the old
  name are left alone, and a `LET` name that would capture the new name is
  renamed to the first free `<name>_<n>`: with `y` renamed to `x`,
  `=LET(x, 1, x+y)` becomes `=LET(x_1, 1, x_1+x)`.

### Patch Changes

- Recover blocked spills automatically. An anchor showing `#SPILL!` now spills
  as soon as its area is free: when the blocking content is cleared, including
  a blocker that depends on the anchor itself, when another spill that held the
  area goes away, and after a warm snapshot restore. Before, the anchor stayed
  `#SPILL!` until its formula was entered again.
- Report spilled cells inside a range at their real position. With `A1` holding
  `x` and `A2` holding `=SEQUENCE(3)`, `MATCH(2, A1:A5, 0)` returned 2 instead
  of 3 and `UNIQUE(A1:A5)` lost a value; `UNIQUE`, `MATCH`, `XLOOKUP` and other
  functions over such ranges are now correct.
- Stop applying element mappings twice when an aggregate reads a computed
  array: `SUM(SEQUENCE(4)/2)` returned 2.5 instead of 5 and `SUM(-SEQUENCE(3))`
  returned 6 instead of -6.
- Keep warm snapshots consistent. A stored spill now records the cells it was
  read from, so lazy spills such as `=A1:A2*1`, `=-A1:A2` and `TOCOL(A1:B2)`,
  and formulas built on them, recalculate after a restore and an edit instead of
  keeping their saved values. A node is stored only together with everything it
  references; one that cannot be, such as an unbounded spill, is left out with
  its dependents and recomputed on demand. A snapshot with unknown node ids or
  nodes that fail to load restores cold with a `[FormulaEngine snapshot]`
  warning, instead of dropping the broken edges silently. An error address
  outside the snapshot no longer makes `serializeEngine` throw.
- Pair `XLOOKUP`'s lookup and return arrays by position. A blank cell in only
  one of the ranges, as in `XLOOKUP(x, A:A, B:B)` with gaps, made every lookup
  `#VALUE!`. An error in `if_not_found` now shows only when nothing matches.

## 0.3.7

### Patch Changes

- Reuse trusted raw-range position indexes and shared range dependencies for
  empty-criteria aggregates instead of resolving every cell for every consumer.
  Preserve dense missing-cell counting, error handling and mutation invalidation;
  retain per-cell callbacks for transformed spills and ranges with spill/cycle
  contributions. Add cold-cache operation-count and semantic regressions.

## 0.3.6

### Patch Changes

- Add `FormulaEngine.clearCache()` for lazy recalculation without clearing workbook data, replacing engine/workbook objects, or changing undo/redo history. Invalidate derived formula, criteria-range, evaluation-order, SCC and spill state and notify existing update subscribers.

## 0.3.5

### Patch Changes

- Replace repeated per-entry transitive dependency searches during range
  evaluation with one iterative candidate-closure/reverse-reachability pass.
  Shared dependency graphs are visited once per range evaluation while preserving
  circular aggregate/spill semantics, including cycles and deep chains.
- Document the compact snapshot size versus cold-recalculation tradeoff; prefer
  warm snapshots for interactive session recovery when reopen latency matters.

## 0.3.4

### Patch Changes

- Index direct/conditional styles and cell data types by workbook, sheet and
  rectangle; cache resolved direct styles with bounded retention and mutation,
  history, reset and rename invalidation. Keep conditional results dynamic.
- Reuse multi-criteria aggregate coordinate indexes for current raw-range results,
  including ranges restored from existing warm snapshots, without losing
  dependencies or mixing transformed ranges, sheets or engines. Preserve finite
  transformed spill values when saving warm state instead of treating their
  source-range metadata as raw values.
- Fix retained virtual formula dependency tracking so conditional styles and
  ad-hoc formula evaluation update after dependent edits, including undo/redo.
- Add compact canonical persistence via
  `serializeEngine({ includeEvaluationState: false })`, preserving all workbook
  data while omitting rebuildable dependency/cache state. The existing loader
  accepts compact saves and existing version 5/6 warm snapshots.
- Add optional `CellStyle.fontFamily` for font fidelity in renderers/importers.

## 0.3.3

### Patch Changes

- Fix read-order-dependent nested aggregates. Shared unresolved dependencies now
  retain their resolution status across dependency graph branches rather than
  allowing a dependent cell to cache a provisional value.
- Scope range expression caches to their workbook and sheet before evaluating
  aggregates. Identical formulas in different workbooks or sheets no longer
  reuse another context's result. Preserve explicit workbook references when
  reducing single-cell ranges to references.

## 0.3.2

### Patch Changes

- Fix stale spilled values when a spilling formula's inputs change.

  A spilling formula is held by two nodes: a spill-meta node that owns the spilled
  result, and the anchor cell's own node that exposes the top-left value. Two
  dependency edges were missing between them, and because invalidation walks
  reverse dependency edges, a node recording no dependencies could never be
  invalidated.

  The spill-meta node never recorded the formula's inputs. When a spill is set up
  through the anchor, the formula is evaluated with the anchor as the context node,
  so the anchor collected the inputs and the spill-meta node was handed only the
  result. Separately, the anchor dropped its own edge to the spill-meta node on
  re-evaluation, because that path clears the collected dependencies first and
  never re-added the edge.

  The effect was that results depended on read order. Reading a spill member first
  happened to build the missing edge, so `C1 =SEQUENCE(3,1,A1)` recalculated
  correctly; reading the anchor first left both nodes with no dependencies, so
  editing `A1` never invalidated them and every later read returned a stale cached
  value for the lifetime of the engine. Both nodes now record the edges they
  depend on, so a spilling formula recalculates the same way regardless of read
  order.

## 0.3.1

### Patch Changes

- Add the `UNIQUE` function and a `#CALC!` error value.

  `UNIQUE(array, [by_col], [exactly_once])` returns the distinct rows of a range or
  array, in order of first occurrence. `by_col` compares columns instead of rows,
  and also picks the orientation of a 1-D input. `exactly_once` keeps only the
  rows or columns that occur a single time. Both flags accept booleans, or `1`
  and `0`.

  Values must share a type to be equal, so `1`, `"1"` and `TRUE` stay distinct,
  while text compares case-insensitively and keeps the casing of its first
  occurrence. `array` may be unbounded along the deduplicated dimension, since the
  tail of empty rows or columns collapses into a single empty entry; unbounded
  along the other dimension is `#VALUE!`, because that dimension sizes the result.

  An error inside `array` propagates. When `exactly_once` removes every row,
  `UNIQUE` returns the new `FormulaError.CALC` (`#CALC!`) value.

## 0.3.0

### Minor Changes

- 429a5f9: Add `WorkbookData` bulk import and change `addWorkbook` to an options object.

  `addWorkbook` now takes `{ workbookName, data? }` instead of a bare string. With
  `data`, a whole workbook — sheets and content, tables, named expressions, cell
  and conditional styles, cell data types, range metadata, and cell/sheet/workbook
  metadata — is applied in one undo/redo step, in dependency order so table
  headers resolve against the content written for them.

  Areas inside `WorkbookData` are sheet-scoped rather than workbook-scoped, so the
  same data can be imported under any workbook name. `content` and `cellMetadata`
  accept a `Map` or a plain object, letting the shape survive a JSON round trip.

  Migration: `engine.addWorkbook("Book")` becomes
  `engine.addWorkbook({ workbookName: "Book" })`.

## 0.2.28

### Patch Changes

- Add structured, consumer-safe cell error details with an optional concrete
  source address for error navigation.

## 0.2.27

### Patch Changes

- Prevent function-like words and valid punctuation inside structured-reference
  columns and workbook qualifiers from producing false syntax diagnostics.

## 0.2.26

### Patch Changes

- Add safe compact and pretty formula formatting APIs, normalize valid formula
  cell content on write, and suppress updates and undo entries for formatting-only
  edits while preserving plain text and invalid formulas exactly.

## 0.2.25

### Patch Changes

- Add an immutable, editor-neutral catalog of built-in functions with aliases,
  categories, signatures, descriptions, and structured parameter metadata.

## 0.2.24

### Patch Changes

- Add editor-neutral formula analysis with semantic tokens, tolerant diagnostics,
  caret reference lookup, and workbook-aware cell, range, table, name, and 3D
  reference targets.

## 0.2.23

### Patch Changes

- Add full-fidelity sheet cloning with formula rewriting, table duplication, and undo/redo support.

## 0.2.22

### Patch Changes

- Add identity-aware workbook and sheet rename and deletion subscriptions, including transaction and undo/redo support.

## 0.2.21

### Patch Changes

- Preserve existing cell contents when tables expand into populated rows or columns, while rejecting table-to-table overlaps.
- Move contained table definitions along with cut-pasted ranges, including infinite-row tables moved by open-ended selections.

## 0.2.20

### Patch Changes

- Restore the version 6 engine implementation, including cell data types,
  incremental undo/redo history, mutation observers, and version 5 snapshot
  migration, while retaining the updated documentation and demo site.

## 0.2.19

### Patch Changes

- 63084eb: Implement approximate `MATCH` modes for ascending and descending lookup arrays.

## 0.2.18

### Patch Changes

- Preserve existing style properties when overlapping cell style rules contain undefined values, including after clearing part of a styled range.

## 0.2.17

### Patch Changes

- Implemented incremental, memory-bounded undo/redo

## Unreleased

### Changed

- Replace full-engine undo/redo snapshots with incremental reversible history
  entries, bounded by both entry count and estimated retained bytes.
- Rename the undo/redo `maxDepth` option and state field to `maxEntries`, add
  `maxBytes`, and report `undoBytes` and `redoBytes` in history state.
- Stream large workbook deltas in bounded fragments, preserve ordered
  collection state exactly, and use a history barrier for metadata whose
  reachable memory cannot be safely cloned and measured.
- Make synchronous transactions atomic against the history byte budget, and
  keep in-place sparse edits, imports, and bulk append/delete replay from
  repeatedly rebuilding unrelated workbook state.

## 0.2.16

### Patch Changes

- Add range-based General, Text, Number, and Boolean cell data types with snapshot, undo/redo, copy/paste, fill, and autofill support.
- 7590825: Keep table column definitions and structured references synchronized when headers change, populate blank headers with unique default names, and reject duplicate table column names.

## 0.2.15

### Patch Changes

- add undo/redo

## 0.2.14

### Patch Changes

- improve search performance

## 0.2.13

### Patch Changes

- Improve cell styling

## 0.2.12

### Patch Changes

- add range metadata + add wrapText style prop

## 0.2.11

### Patch Changes

- Fix serialized engine restores so named-expression scopes are rebuilt from workbook snapshots and orphan scopes are ignored.
- fix bug with serialization / deserialization

## 0.2.10

### Patch Changes

- add search and replace

## 0.2.9

### Patch Changes

- remove the schema api this will become a user responsibility

## 0.2.8

### Patch Changes

- add removeSchema support to the engine

## 0.2.7

### Patch Changes

- better cache invalidation

## 0.2.6

### Patch Changes

- bump snapshot version

## 0.2.5

### Patch Changes

- Fixed a dependency-resolution bug where open ranges could keep stable frontier candidates alive for too long during reevaluation.

  In the failing cases, a formula would reference an open range such as `Q18:Q` or `A1:INFINITY`. The range planner correctly added frontier candidates that might spill into that range, but once those candidates had been evaluated and were known not to spill, they were not being discarded early enough. That left avoidable transient dependencies in the graph, caused repeated evaluation-plan rebuilds, and made some operations such as toggling a table feel much slower than they should.

  This change lets range nodes prune frontier candidates as soon as those candidates are stable enough to prove that they do not spill into the range, and it excludes discarded frontier dependencies from traversal of the active dependency graph. In practice this removes a large amount of false-positive reevaluation work for open ranges and significantly improves recalculation time after table and range-shape changes.

  While fixing that, the reevaluation path for circular open-range cases was also tightened so a formula such as `SUM(A1:INFINITY)` does not count its own circular result back into the aggregate on a later rerun.

## 0.2.4

### Patch Changes

- sheet rename fixes

## 0.2.3

### Patch Changes

- snapshot whole engine + performant edits

## 0.2.2

### Patch Changes

- add addGridSchema method to the engine class

## 0.2.1

### Patch Changes

- fix type issue

## 0.2.0

### Minor Changes

- ### Added

  - **Grid Schema Support**: Added `addGridSchema()` method to define schemas for 2D ranges of cells. Grid schemas provide readonly `columns` and `rows` getters for column-major and row-major array access, plus `setValue()` and `getValue()` methods for individual cell access.

  ### Changed

  - **Explicit Write Functions**: Schema definitions now use explicit `write` functions instead of guessing how to serialize parsed values. This provides better type safety and control over serialization.
  - **Optional Write for Primitive Types**: The `write` function is now optional when `parse` returns a `SerializedCellValue` (number, string, or boolean). For complex types (objects), `write` is required and TypeScript will error if omitted.
  - **GridOrm API**: Removed `columns` and `rows` setters from GridOrm. Use `setValue(value, position)` instead. The getters now return readonly arrays.
  - **Table Headers**: Updated `defineHeader()` helper to support optional `write` functions. Headers now use explicit write functions for serialization.

  ### Migration Guide

  **Grid Schemas:**

  ```typescript
  // Before: No grid schema support
  // After:
  .addGridSchema(
    "matrix",
    { workbookName: "wb", sheetName: "s" },
    { start: { col: 0, row: 0 }, end: { col: 9, row: 9 } },
    (value) => parseNumber(value)
    // write optional for primitive types
  )

  // Access grid data
  schema.matrix.columns  // readonly number[][]
  schema.matrix.rows     // readonly number[][]
  schema.matrix.setValue(42, { col: 0, row: 0 })
  schema.matrix.getValue({ col: 0, row: 0 })
  ```

  **Cell Schemas with Complex Types:**

  ```typescript
  // Before: Guessing serialization
  .addCellSchema("item", addr, (v) => ({ value: v, meta: true }))

  // After: Explicit write function required
  .addCellSchema(
    "item",
    addr,
    (v) => ({ value: v, meta: true }),
    (item) => ({ value: item.value })  // REQUIRED for object types
  )
  ```

  **Table Headers:**

  ```typescript
  // Before: Simple header definition
  { value: { parse: parseNumber, index: 0 } }

  // After: Use defineHeader helper
  { value: defineHeader(0, parseNumber) }
  // Or with custom write:
  { value: defineHeader(0, parseNumber, (v) => ({ value: v })) }
  ```

### Patch Changes

- export defineHeader

## 0.1.0

### Minor Changes

- ### Added

  - **Grid Schema Support**: Added `addGridSchema()` method to define schemas for 2D ranges of cells. Grid schemas provide readonly `columns` and `rows` getters for column-major and row-major array access, plus `setValue()` and `getValue()` methods for individual cell access.

  ### Changed

  - **Explicit Write Functions**: Schema definitions now use explicit `write` functions instead of guessing how to serialize parsed values. This provides better type safety and control over serialization.
  - **Optional Write for Primitive Types**: The `write` function is now optional when `parse` returns a `SerializedCellValue` (number, string, or boolean). For complex types (objects), `write` is required and TypeScript will error if omitted.
  - **GridOrm API**: Removed `columns` and `rows` setters from GridOrm. Use `setValue(value, position)` instead. The getters now return readonly arrays.
  - **Table Headers**: Updated `createHeader()` helper to support optional `write` functions. Headers now use explicit write functions for serialization.

  ### Migration Guide

  **Grid Schemas:**

  ```typescript
  // Before: No grid schema support
  // After:
  .addGridSchema(
    "matrix",
    { workbookName: "wb", sheetName: "s" },
    { start: { col: 0, row: 0 }, end: { col: 9, row: 9 } },
    (value) => parseNumber(value)
    // write optional for primitive types
  )

  // Access grid data
  schema.matrix.columns  // readonly number[][]
  schema.matrix.rows     // readonly number[][]
  schema.matrix.setValue(42, { col: 0, row: 0 })
  schema.matrix.getValue({ col: 0, row: 0 })
  ```

  **Cell Schemas with Complex Types:**

  ```typescript
  // Before: Guessing serialization
  .addCellSchema("item", addr, (v) => ({ value: v, meta: true }))

  // After: Explicit write function required
  .addCellSchema(
    "item",
    addr,
    (v) => ({ value: v, meta: true }),
    (item) => ({ value: item.value })  // REQUIRED for object types
  )
  ```

  **Table Headers:**

  ```typescript
  // Before: Simple header definition
  { value: { parse: parseNumber, index: 0 } }

  // After: Use createHeader helper
  { value: createHeader(0, parseNumber) }
  // Or with custom write:
  { value: createHeader(0, parseNumber, (v) => ({ value: v })) }
  ```

## 0.0.22

### Patch Changes

- Rename API to Schema throughout the codebase. All API-related types, functions, and properties have been renamed to use "schema" terminology. `defineApi` is now `defineSchema`, `engine.api` is now `engine.schema`, and all related types have been updated. The `createApi` alias has been removed. Added runtime schema methods `addTableSchema` and `addCellSchema` to FormulaEngine for dynamically registering schemas after engine creation. Schema validation now works correctly for both initial and runtime-added schemas.

## 0.0.21

### Patch Changes

- Remove custom methods layer from API. The `methods` parameter has been removed from `addTableApi` and `addCellApi`. ORM methods (`findWhere`, `findAllWhere`, `append`, `updateWhere`, `removeWhere`, `count` for tables and `read`, `write`, `getAddress` for cells) are now exposed directly on the API namespace.

## 0.0.20

### Patch Changes

- Implement Excel-like table spill restrictions: spilling formulas inside tables and formulas that would spill into table areas now return #SPILL! error instead of succeeding.

## 0.0.19

### Patch Changes

- fix exported lib

## 0.0.18

### Patch Changes

- api support and new state management

## 0.0.17

### Patch Changes

- - Add cut-paste reference updates, multi-area style support, and proper overlapping range handling

    **Breaking Changes:**

    - `ConditionalStyle` and `DirectCellStyle` now use `areas: RangeAddress[]` instead of `area: RangeAddress`
      - Enables Excel-like multi-area styles when cutting cells from styled regions
      - Cutting a cell from a styled area now punches a hole, creating multiple areas in the same style object

    **New Features:**

    - Cell and range references are automatically updated when cells are cut and pasted
      - When a cell (e.g., A1) is cut to a new location (e.g., D5), all formulas referencing A1 update to D5
      - Both relative and absolute references are updated (e.g., `=$A$1` becomes `=$D$5`)
      - Range references update only when the entire range is moved (e.g., `=SUM(A1:B2)` → `=SUM(D5:E6)`)
      - Partial range moves don't update references (prevents breaking formulas)
      - Works across sheets and workbooks
    - Added `moveCell(source, target)` method - programmatically move cells with reference updates
    - Added `moveRange(sourceRange, target)` method - programmatically move ranges with reference updates
    - Added `getAllCellStyles()` and `getAllConditionalStyles()` methods to engine for testing/serialization
    - Cut operations always use `pasteCells()` not `fillAreas()` (Excel behavior)

    **Style System Improvements:**

    - **Multi-area style support**: Styles can now cover non-contiguous ranges
      - Cutting cell L11 from styled area J3:N20 creates 4 areas: J3:N10, J12:N20, J11:K11, M11:N11
      - Hole-punching keeps all areas in the same style object (not separate styles)
    - `clearCellStylesInRange()` now properly punches holes instead of creating new style objects
    - All style operations (filtering, updating, rendering) now work with multiple areas

    **Architecture Improvements:**

    - Restructured cut flow: snapshot → remove → update references → apply (proper Excel-like behavior)
    - Renamed internal `copyCells()` to `pasteCells()` in CopyManager for consistency
    - Enhanced `CellSnapshot` to include style information for proper cut operations
    - Added `WorkbookManager.updateFormulasExcluding()` to update references excluding moved cells
    - All sheet modifications now go through proper WorkbookManager APIs (maintains indexes)
    - Snapshot-based copying handles overlapping source/target ranges correctly
    - Formulas in moved cells are adjusted once; formulas referencing moved cells updated separately

    **Bug Fixes:**

    - Fixed overlapping cut operations (e.g., cutting A1:B4 to B1:C4 now correctly clears A1-A4)
    - Fixed cache invalidation when modifying cells during cut operations
    - Cut cells now properly carry their styles to the new location

    **Implementation Details:**

    - New `src/core/cell-mover.ts` module with formula transformation utilities
    - Split CopyManager into `cutCells()` (move flow) and `copyOnlyCells()` (copy flow)
    - Cut and move operations use identical underlying logic
    - Comprehensive test coverage: 563 tests passing, including overlapping ranges and multi-area styles

## 0.0.16

### Patch Changes

- fix clone workbook

## 0.0.15

### Patch Changes

- improve color management

## 0.0.14

### Patch Changes

- 7c967ac: add smartPaste

## 0.0.13

### Patch Changes

- enable copy of value only or style only

## 0.0.12

### Patch Changes

- add getStyleForRange

## 0.0.11

### Patch Changes

- make getCellStyles take a range

## 0.0.10

### Patch Changes

- Update range/copy-manager logic and tests

## 0.0.9

### Patch Changes

- add engine.clearCellStyles()

## 0.0.8

### Patch Changes

- add copyCells method and expand styling capabilities

## 0.0.7

### Patch Changes

- export color utils

## 0.0.6

### Patch Changes

- add conditional formatting and cell styling support

## 0.0.5

### Patch Changes

- remove last index.ts file

## 0.0.4

### Patch Changes

- fix base import path

## 0.0.3

### Patch Changes

- export utils

## 0.0.2

### Patch Changes

- add lib entry

## 0.0.1

### Patch Changes

- initial release
