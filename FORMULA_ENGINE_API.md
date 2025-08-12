In the UI I interface with the formula-engine like this:
1. add an empty sheet
2. setSheetContent - to set initial sheet content (which is a Map<string, boolean | number | string> where the values of the Map can include strings which are formulas)
3. getSheetContent is called to get the serialized data (the data passed to setSheetContent)
4. The UI makes an initial render with the data from getSheetContent. From the UI each cell will call getCellValue to get the calculated value for each cell which is displayed to the user
5. The user double clicks a cell to update it. He will now not see the value of the cell (from getCellValue) but the "content" i.e. the formula from getSheetContent (unless it is a spilled cell, which is handled a bit different in the UI) in an input box.
6. The user updates the cell by changing a formula, and setSheetContent is called again with the updated spreadsheet content
7. If the sheet content is updated, then listeners for the sheet, from listenToSheetContentUpdate("Sheet1, callback) are triggered.
8. The UI has a listener for sheetContentUpdates, and when called the UI is re-rendered
9. getSheetContent is called to get the serialized data (the data passed to setSheetContent) (step 3)
10. The UI will re-render with the data from getSheetContent. From the UI each cell will call getCellValue to get the calculated value for each cell which is displayed to the user. getCellValue will return a cached value if available, or re-calculate if dirty








* API surface: addSheet, setSheetContent, getSheetContent, getCellValue, listenToSheetContentUpdate
* Keep a dependency DAG with both forward (precedents) and reverse (dependents) edges.
* Maintain a dirty set and only reevaluate nodes whose inputs actually changed.
* Spill results are modeled explicitly so that membership changes (shape/placement) propagate correctly.
* Evaluate in topological order
* Evaluation model: Lazy and reader-driven. Each getCellValue() evaluates exactly what it needs; internal caches avoid redoing work on subsequent calls to getCellValue()
* Identity: Sheets, tables, and named expressions are keyed by user-provided names (case-insensitive recommended). Renames are strong renames (rewrite formulas).
* Spills: All-or-nothing. If any target cell is blocked/out-of-bounds → origin returns #SPILL!; no members are created.



### setSheetContent:
1. **Parse first, don’t evaluate yet**

   * For each edited cell:
     * Parse to `ast` (or `empty`/`literal`).
     * Extract *static* references (cells, ranges, named expressions, tables). Don’t resolve dynamic ranges yet; just record their definitions.
   * Keep `oldDeps` (from graph) and `newDeps` (from parse).

2. **Compute a graph delta**

   * For each edited cell `c`:

     * `removed = oldDeps(c) - newDeps(c)`
     * `added   = newDeps(c) - oldDeps(c)`
   * Build changes:

     * Remove edges `c -> removed`, update reverse edges.
     * Add edges `c -> added`, update reverse edges.
   * If `c` previously had a `SpillGroup`, mark all its old members as **dirty** and clear the alias map entries (they may stop being members).
   * Mark `c` as **dirty** (content changed).

3. **Seed the dirty frontier**

   * `dirty = edited cells ∪ prior spill members that got invalidated`.
   * You do **not** need to add transitive dependents yet; we’ll discover them while scheduling.

4. **Induce the affected subgraph**

   * BFS/DFS via `revDeps` from the current `dirty` to collect all **potential** recompute targets `affected`.
   * (Optional optimization) Stop walking past nodes whose inputs’ `version` didn’t change once known — see step 7. But a first pass can just include them all.

5. **Incremental topological order over the affected set**

   * Use Kahn’s algorithm restricted to `affected`:

     * Compute each node’s `inDegreeAffected` = number of deps within `affected`.
     * Initialize a queue with nodes where `inDegreeAffected == 0`.
     * Pop, emit, decrement successors.
   * **Cycle detection**: if you can’t consume all nodes, you’ve got a cycle within `affected`. Mark those nodes with a `#CYCLE!` error and treat their dependents as seeing an error value. (If you support iterative calc, this branch differs.)

6. **Evaluate in topo order; handle spill mechanics carefully**
   For each node `n` in order:

   * If `contentKind` is `literal` or `empty`:

     * `newValue = literal or EMPTY`
   * If `formula`:

     * Evaluate with current values of precedents (which are already evaluated due to topo order).
     * If the formula **spills** (e.g., `SEQUENCE`, dynamic arrays):

       * Compute the **shape** and **values**.
       * Compute **target members** (origin’s anchor + shape).
       * Check **obstructions**; if obstructed, set `#SPILL!` on origin, and *don’t* write members.
       * Otherwise:

         * Build a new `SpillGroup` and member set.
         * Update the alias map: for each member cell `m`, `spillOwner[m] = origin`.
         * For cells that were members before but not now: mark as dirty (they’ll become empty or take on old content again).
         * **Write semantics**: you can either:

           1. Store the full 2D array on the origin and treat members as *read-through aliases* to slices of the origin (cleanest dependency graph), or
           2. Materialize values into member `Node.value` and give each member a single dep on origin (more edges but simpler reads).
              Approach (1) is usually faster and avoids churn in the graph.
   * **Change detection**:

     * Compare `newValue` to `oldValue` (deep equality for arrays; treat `NaN`s carefully).
     * If changed: set `Node.value = newValue`; bump `Node.version++`; add all `revDeps(n)` to a `nextDirty` set.
     * If unchanged: leave version; do **not** enqueue dependents.

   After finishing pass, set `dirty = nextDirty`, **recompute topo just over `dirty`** and repeat until `dirty` is empty. (In practice you’ll fold this into a single pass by pushing dependents when a value changes.)

