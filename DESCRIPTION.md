## FormulaEngine — Project Description
FormulaEngine is a TypeScript formula engine for spreadsheet-like computation with strong Excel compatibility.
It is headless, embeddable, and optimized for sparse data and dynamic arrays with spill semantics.
The evaluation model is lazy and viewport-driven: compute only what the UI or explicit readers demand.
A global dependency graph (DAG) is maintained eagerly on edits, storing forward deps and reverse revDeps.
Nodes represent concrete cells, materialized spill members, named expressions, and watcher-backed ranges.
Each cell node tracks an address, kind, AST (for formulas), deps, revDeps, value, versions, and dirtiness.
Change propagation uses versioned values and input fingerprints to avoid unnecessary recomputation.
Spills are modeled with materialized members: each member depends solely on its origin cell.
Spill origins define symbolic coverage with RectUnions that may include infinite ends.
Realized spill area is coverage truncated by nearest obstructions and sheet boundaries.
Infinite and finite spills materialize members on demand only when visible or referenced.
If the origin’s anchor is obstructed, the origin yields #SPILL! without realizing members.
Editing obstructions triggers re-evaluation and possible spill growth or shrink.
The engine maintains per-origin coverage indices and nearest-blocker indices for fast truncation math.
Multi-sheet references are first-class through SheetRange and MultiRange abstractions.
3D ranges bind to sheet index order, staying correct across renames and reorders.
Aggregators over 2D and 3D ranges depend on watchers, not on every individual cell.
The parser builds ASTs directly using a recursive descent design with context-sensitive references.
References can be single cells, ranges, named expressions, table areas, and dynamic ranges.
Dynamic ranges (OFFSET, INDEX, FILTER, SORT) are stored symbolically and resolved at evaluation time.
Volatile functions advance a global epoch and only bump nodes whose values actually change.
Cycle detection marks strongly connected components with #CYCLE! during restricted topo sorts.
Equality rules follow JS with NaN equals NaN and optional +0 versus -0 distinctions.
Array formulas support NumPy-style broadcasting and vectorized operations.
Origins expose a tile API for sliceable evaluation, avoiding full-array materialization.
Tiles are cached per origin with independent versions for incremental recomputation.
Member nodes remember source tile versions to skip unnecessary bumps.
Named expressions are first-class with workbook or sheet scope and stable nameIds.
Named formulas behave like hidden compute nodes and may be relative to a caller anchor.
Named regions compile to range watchers (2D or 3D) and recompile on definition changes.
Scope shadowing lets sheet-scoped names override workbook names consistently.
Table support introduces a registry of tables, columns, filters, sort, and calculated columns.
Structured references map to table areas using stable tableIds and columnIds.
Table watchers compile to per-sheet rectangles and mark subscribers dirty on structure changes.
Sorting emits structure events; order-insensitive aggregations may not recompute.
Table cells are spill obstructions; adjacent spills truncate at table boundaries.
Calculated columns use relative row context via [@Column] semantics.
The lazy evaluation pipeline separates edit-time graph maintenance from render-time computation.
On edit, formulas parse, deps delta is applied, watchers are registered, and nodes mark dirty.
On render, the engine picks origins intersecting the viewport and spill-pending origins.
The prerequisite closure is evaluated across sheets using fingerprints to short-circuit work.
Spill membership is reconciled, adding or removing materialized members as needed.
The demand set includes viewport cells, pinned cells, and upstream prerequisites via concrete deps.
A restricted topological sort drives evaluation, and results render with final values.
Garbage collection dematerializes offscreen and unreferenced spill members to bound memory.
Events are buffered per calculation cycle and flushed once with final deduped values.
Subscribers can observe specific cells, entire sheets, and optional spill membership changes.
Subscribing to a cell pins it so it remains fresh offscreen without forcing whole-sheet evaluation.
Emission guarantees at-most-once-per-cell per cycle and sheet-batched delivery.
Causes tagged on events include userEdit, recalc, volatile, spillResize, and structure.
The API exposes getters for values, formulas, and serialized content per cell, sheet, and workbook.
Data manipulation includes setCellContent, setSheetContent (Map-friendly), and range operations.
Sheet management covers add/remove/rename, existence checks, and count with stable sheetIds.
Dependency analysis provides getCellPrecedents and getCellDependents for cells and ranges.
Evaluation control supports suspend/resume and formula utilities like normalize and validate.
Undo/redo follows a command pattern with proper state tracking and stack management.
React integration offers hooks for subscribing to cells and sheets without over-evaluation.
The internal storage is sparse-aware; only populated cells consume memory and processing.
Maps and address strings provide O(1) access and efficient iteration over populated cells.
Types are strict; values are number, string, boolean, error, or undefined for empty.
Errors use spreadsheet codes such as #DIV/0!, #N/A, #NAME?, #NUM!, #REF!, #VALUE!, #CYCLE!, #ERROR!.
Additional table error codes include #TABLE!, #COLUMN!, and #SELECTOR! when enabled.
The function library is modular and extensible via a type-safe registry and plugins.
Math, logical, text, lookup, array, and info functions follow Excel behavior where practical.
Broadcasting rules implement scalar-to-array and compatible array operations with spilling.
Lookup features include VLOOKUP, INDEX, MATCH, XLOOKUP, and INDIRECT semantics.
Address utilities parse A1 references, ranges, and structured references reliably.
The dependency graph stores edges for concrete deps and watcher tokens for ranges and tables.
Watchers live in spatial indices per sheet and a manager for 3D expansions over sheet ranges.
Workbook metadata uses stable sheetIds with byId and byName maps synced to renames.
Sheet range semantics follow index order to keep 3D references stable across moves and inserts.
Performance focuses on minimal recomputation, tile-level incrementality, and cache reuse.
Fingerprints are derived from dependency valueVersions and configurable hash functions.
Configuration includes tile size, GC thresholds, fingerprint hash, and obstruction policies.
Spill reconciliation diffs previous and new realized regions to add and remove members precisely.
Aggregators over infinite or divergent sources return #NUM! per defined semantics.
Address-sensitive functions work naturally because members are real cells with addresses.
The engine emits structure events when names, tables, or spills cause layout-level changes.
Testing covers unit, integration, property-based, performance, and Excel compatibility cases.
Viewport tests confirm scrolling only evaluates demanded cells and origins.
Spill tests ensure shape changes do minimal work and do not bump unaffected members.
Infinite spill tests validate truncation and on-demand materialization under blockers.
Events tests prove per-cycle batching, dedupe, and correct causes and versions.
Multi-sheet tests exercise 3D ranges, sheet index changes, and dynamic sheet references.
Named expression tests verify scope shadowing, relative names, and region recompilation.
Table tests validate column renames, resize, filters, sorts, and calculated columns.
Developer experience emphasizes clear TypeScript APIs and exhaustive IntelliSense.
The engine is designed for Bun-based workflows for dev, test, and builds.
Public examples demonstrate spreadsheets, dynamic arrays, and dependency analysis in React.
The architecture separates parsing, dependency management, and evaluation for maintainability.
Parsing constructs ASTs once and reuses them via relative addressing and caching.
Dependency graph updates are incremental, keeping the DAG stable and cheap to mutate.
Evaluation is demand-driven with strict version bump rules to prevent cascading churn.
Memory is bounded via sparse structures, tile caches, and GC of offscreen members.
Extensibility supports custom functions, plugins, and pluggable parsers if needed.
Serialization focuses on stable identities (nameId, sheetId, tableId, columnId) over display names.
Error handling is comprehensive and recoverable via IFERROR/IFNA and consistent propagation.
The end goal is sub-second recalculation for large workbooks with excellent developer ergonomics.
This DESCRIPTION.md is the canonical high-signal overview used as context for future prompts.