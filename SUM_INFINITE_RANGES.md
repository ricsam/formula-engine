# Technical Spec: “Frontier spill candidates” for open-ended SUM rectangles

## Goal

Given:

* A **rectangular target region** `R = [rows rT..rB] × [cols cL..cR]`

  * `rB` and/or `cR` may be **open-ended** (unbounded below/right).
* A **finite set of known non-empty cells** `O` (values or formulas).
* The subset of **formula cells** `F ⊆ O`.

Return a set `C ⊆ F` of **formula cells outside R** that **must be evaluated** *before* computing an aggregate (e.g., `SUM`) over `R`, because they may **spill into R** (down and/or right). The idea is: only formula cells on the **immediate frontier** that could reach R without being blocked by closer non-empty cells matter. Farther formula cells in the same “approach line” would error on spill (blocked), so you can safely skip them.

> Assumption on spill directions
> Spills are rectangular and extend **only down and/or right** from the formula’s cell (Excel-style). Therefore, only sources **above** `R` (for downward reach) and **left of** `R` (for rightward reach) can enter `R`.
> If your engine later supports up/left spills, apply the same logic symmetrically.

## Inputs

* `rT, rB, cL, cR` with `rB = +∞` and/or `cR = +∞` allowed.
* `O`: finite set of non-empty cells `(r, c)`, each tagged as **value** or **formula**.
* `F`: the formula subset of `O`.

## Output

* `C`: minimal set of **frontier formula cells** whose evaluation (incl. transitive deps) is necessary to determine the aggregate over `R`.

## Definitions

* A cell **blocks** a spill rectangle if it is **non-empty** (`∈ O`) and lies inside that would-be rectangle. (This is spill error semantics.)
* For a fixed **column** `c`, define `Above(c)` as non-empty cells with that column and row `< rT`.
  For a fixed **row** `r`, define `Left(r)` as non-empty cells with that row and column `< cL`.
* **Nearest-above non-empty** in column `c`: the cell in `Above(c)` with **maximum row** `< rT` (closest to `rT`).
* **Nearest-left non-empty** in row `r`: the cell in `Left(r)` with **maximum col** `< cL` (closest to `cL`).

## Algorithm (deterministic, finite)

### 0) Preliminaries

* You never need to search beyond the **finite** columns/rows that actually appear in `O`.
* When `cR = +∞`, consider **all columns `c ≥ cL` that appear in `O`**, plus `cL` itself even if currently empty.
  When `rB = +∞`, consider **all rows `r ≥ rT` that appear in `O`**, plus `rT` itself.
* Having only empties past the last known column/row cannot create new candidates because candidates must be formula cells (which live in `F ⊆ O`, hence finite).

### 1) Top frontier (potential downward spills)

For every column `c` that **intersects the rect horizontally**:

* If `cR` is finite: `c ∈ [cL..cR]`.
* If `cR = +∞`: `c ∈ {c | c ≥ cL and ∃(r, c) ∈ O}` (i.e., all known columns at/after `cL`).

Do:

1. Let `A = { (r, c) ∈ O | r < rT }`.
2. If `A` is empty → no candidate from column `c`.
3. Let `(r*, c)` be the **nearest-above non-empty** (max `r` in `A`).
4. If `(r*, c) ∈ F` (formula), **add** it to `C`; else (value) it **blocks** and **no** candidate from this column.

> Rationale: Any formula farther above (`r < r*`) would spill into `(r*, c)` first and error, so it cannot reach `R`.

### 2) Left frontier (potential rightward spills)

For every row `r` that **intersects the rect vertically**:

* If `rB` is finite: `r ∈ [rT..rB]`.
* If `rB = +∞`: `r ∈ {r | r ≥ rT and ∃(r, c) ∈ O}`.

Do:

1. Let `L = { (r, c) ∈ O | c < cL }`.
2. If `L` is empty → no candidate from row `r`.
3. Let `(r, c*)` be the **nearest-left non-empty** (max `c` in `L`).
4. If `(r, c*) ∈ F` (formula), **add** it to `C`; else (value) it **blocks** and **no** candidate from this row.

> Rationale: Any formula farther left (`c < c*`) would spill into `(r, c*)` first and error, so it cannot reach `R`.

### 3) In-rect formulas

Independently, for the aggregate:

* Evaluate **all formula cells in `R`** and their **transitive dependencies**.

### 4) De-dup & closure

* `C` may include duplicates from steps (1) and (2); **deduplicate**.
* The full evaluation set is:

  * `Eval = (F ∩ R) ∪ C`, and then **transitive closure of dependencies** of every cell in `Eval`.
* If evaluating any candidate produces an **infinite spill** that intersects `R`, the aggregate result may be **+∞** per your engine’s rule.

### 5) Complexity

* Build index by columns and rows from `O` (hash maps or sorted structures).
* Each column/row contributes at most **one** candidate.
* Overall: `O(|O| log |O|)` if sorting per row/column, or `O(|O|)` with per-line running maxima.

## Correctness intuition

* A spill that reaches `R` must cross its **top** boundary (some column) or **left** boundary (some row).
* On that crossing line, the first non-empty cell encountered when walking **up** or **left** from the boundary is the only possible **unblocked** origin cell; anything farther is provably blocked.
* Considering both top and left frontiers covers **diagonal** origins: any diagonal origin that isn’t nearest on at least one crossing line is blocked on that line.

---

# Test Suite (ASCII / Markdown)

Conventions:

* `F` = formula cell (outside R unless noted)
* `V` = value cell (non-empty, blocks)
* `.` = empty
* `[x]` = cell inside the target rect `R` (x may be `.`, `F`, or `V`)
* Coordinates shown for clarity (Cols A..E, Rows numeric)
* Expected result lists:

  * `FrontierTop`: candidates from Step 1
  * `FrontierLeft`: candidates from Step 2
  * `C = FrontierTop ∪ FrontierLeft`
  * Also list `InRectFormulas`

---

### Test 1 — Basic top spill candidate

Rect: `R = B10:D∞` (rows `≥ 10`, cols `B..D`)

```
      A   B   C   D   E
r8    .   .   F   .   .
r9    .   V   .   F   .
r10   .  [.] [.] [.]  .
r11   .  [.] [.] [.]  .
```

* Column B above r10: nearest non-empty at r9 is `V` → blocks → no candidate.
* Column C above r10: nearest non-empty at r8 is `F` → candidate.
* Column D above r10: nearest non-empty at r9 is `F` → candidate.
* Left frontier: rows `≥10` have nothing left of `B`.

**Expected**

* `FrontierTop = { C8, D9 }`
* `FrontierLeft = ∅`
* `C = { C8, D9 }`
* `InRectFormulas = ∅`

---

### Test 2 — Blocking stack (only nearest counts)

Rect: `R = C6:D10` (bounded)

```
      A   B   C   D
r3    .   .   F   .
r4    .   .   F   .
r5    .   .   .   .
r6   [.] [.] [.] [.]
r7   [.] [.] [.] [.]
```

* Column C above r6: nearest non-empty is `C4 = F` → candidate.
  `C3 = F` is *farther* and is **ignored** (would spill into `C4` first → error).
* Column D above r6: none.
* Left frontier: rows 6..10 have nothing to the left of column C.

**Expected**

* `FrontierTop = { C4 }`
* `FrontierLeft = ∅`
* `C = { C4 }`
* `InRectFormulas = ∅`

---

### Test 3 — Left frontier with blocker value

Rect: `R = C5:∞` (rows `≥ 5`, cols `≥ C`)

```
      A   B   C   D   E
r4    .   F   .   .   .
r5   [.] [V] [.] [.]  .
r6   [.] [.] [.] [.]  .
```

* Top frontier: for columns `C..E`, above r5 there’s nothing non-empty → none.
* Left frontier (rows ≥5):

  * Row 5: nearest-left non-empty is `B5 = V` → blocks → no candidate.
  * Row 6: nearest-left non-empty is `B6` (empty) and `A6` (empty) → none.
* Note `B4 = F` is above row 5, **but not in the same row**; it cannot spill right **across** a row boundary to reach R (spills extend within its own rows downward/right). It could only reach via **downward** into row 5 first, but top frontier for column B is irrelevant since column B is **outside** R’s horizontal span.

**Expected**

* `FrontierTop = ∅`
* `FrontierLeft = ∅`
* `C = ∅`
* `InRectFormulas = ∅`

---

### Test 4 — Left frontier candidate (nearest formula wins)

Rect: `R = C5:D8` (bounded)

```
      A   B   C   D   E
r5   [.] [F] [.] [.]  .
r6   [.] [V] [.] [.]  .
r7   [.] [F] [.] [.]  .
r8   [.] [.] [.] [.]  .
```

Left frontier rows 5..8:

* r5: nearest-left non-empty is `B5 = F` → candidate.
* r6: nearest-left non-empty is `B6 = V` → blocks.
* r7: nearest-left non-empty is `B7 = F` → candidate.
* r8: none.

Top frontier: columns C and D above r5 are empty.

**Expected**

* `FrontierLeft = { B5, B7 }`
* `FrontierTop = ∅`
* `C = { B5, B7 }`
* `InRectFormulas = ∅`

---

### Test 5 — Diagonal origin blocked by nearer frontier

Rect: `R = C6:E9` (bounded)

```
      A   B   C   D   E
r4    .   .   .   V   .
r5    F   .   .   .   .
r6   [.] [.] [.] [.] [.]
r7   [.] [.] [.] [.] [.]
```

* Candidate at `A5 = F` (above-left) seems like it could spill into R.
* But column D has a blocker `D4 = V` above `r6`.
  Top frontier for column D → nearest-above non-empty is a **value**, so **no** top candidate there.
* To enter R, the rectangle from `A5` would have to cover `D4` on its way right/down → **blocked**.
  Since `A5` is **not** nearest on any crossing line (neither nearest-above in columns C..E nor nearest-left in rows 6..9), it’s **excluded**.

**Expected**

* `FrontierTop = ∅`
* `FrontierLeft = ∅`
* `C = ∅`
* `InRectFormulas = ∅`

---

### Test 6 — Mixed: top and left frontiers both apply

Rect: `R = B10:D∞` (rows `≥10`, cols `B..D`)

```
      A   B   C   D   E
r7    .   .   F   .   .
r8    F   .   .   .   .
r9    .   .   V   F   .
r10  [.] [.] [F] [.]  .
r11  [.] [.] [.] [.]  .
```

* In-rect formulas: `C10`.
* Top frontier:

  * Column B: none above.
  * Column C: nearest-above is `r9 = V` → blocks.
  * Column D: nearest-above is `r9 = F` → candidate `D9`.
* Left frontier (rows ≥10 seen in `O` are 10 and 11 due to entries):

  * Row 10: nearest-left non-empty is `A10` (empty) / `B10` (empty) → none.
  * Row 11: nearest-left non-empty is `A11` (empty) → none.
* The `A8 = F` is left of the rect but on **row 8**, so it cannot spill right into rows ≥10 without first spilling **down** through rows 9.., but its top approach is blocked by `V` in `C9`.

**Expected**

* `InRectFormulas = { C10 }`
* `FrontierTop = { D9 }`
* `FrontierLeft = ∅`
* `C = { D9 }`

---

### Test 7 — Open-ended right edge

Rect: `R = C3:∞` (rows `≥3`, cols `≥ C`)

```
      A   B   C   D   E   F
r2    .   F   .   .   .   .
r3   [.] [F] [.] [.] [.]  .
r4   [.] [V] [.] [.] [.]  .
```

* Top frontier (columns ≥ C): none above r3 in C..F.
* Left frontier (rows ≥3 that appear): rows 3 and 4.

  * r3: nearest-left non-empty is `B3 = F` → candidate.
  * r4: nearest-left non-empty is `B4 = V` → blocks.
* `B2 = F` is above and left, but not nearest on row 3, and column-wise, it doesn’t help because columns B < C are outside R’s horizontal span; to reach R it must cross row 3 first, where `B3` is nearer and thus blocks.

**Expected**

* `FrontierTop = ∅`
* `FrontierLeft = { B3 }`
* `C = { B3 }`
* `InRectFormulas = ∅`

---

### Test 8 — Infinite spill intersects R

Rect: `R = B5:D∞`

```
      A   B   C   D
r3    .   .   .   .
r4    .   F   .   .
r5   [.] [.] [.] [.]
```

* Top frontier col B: nearest-above non-empty is `B4 = F` → candidate.
* Suppose evaluating `B4` yields an **unbounded** spill downwards (per your engine). Since its rectangle intersects `R`, the aggregate returns **INFINITY**.

**Expected**

* `FrontierTop = { B4 }`
* `FrontierLeft = ∅`
* `C = { B4 }`
* `InRectFormulas = ∅`
* **Aggregate outcome**: `INFINITY` if `B4` evaluates to an infinite spill.

---

### Test 9 — In-rect formulas plus frontier

Rect: `R = B5:C7`

```
      A   B   C   D
r3    .   .   .   .
r4    .   F   .   .
r5   [F] [.] [.]  .
r6   [.] [.] [.]  .
r7   [.] [.] [.]  .
```

* In-rect formulas: `B5`.
* Top frontier:

  * Column B: nearest-above non-empty is `B4 = F` → candidate.
  * Column C: none above r5.
* Left frontier rows 5..7: none to the left of column B.

**Expected**

* `InRectFormulas = { B5 }`
* `FrontierTop = { B4 }`
* `FrontierLeft = ∅`
* `C = { B4 }`

---

### Test 10 — Value walls negate far formulas

Rect: `R = D10:E12`

```
      A   B   C   D   E
r7    F   .   .   .   .
r8    .   V   .   .   .
r9    .   .   V   .   .
r10  [.] [.] [.] [.] [.]
r11  [.] [.] [.] [.] [.]
r12  [.] [.] [.] [.] [.]
```

* Top frontier:

  * Column D: nearest-above non-empty is `C9 = V` (but note: different column; for column D, there is **no** non-empty above) → none.
  * Column E: same → none.
* Left frontier rows 10..12:

  * Nearest-left in each row is empty (no non-empty at `c < D` in those rows) → none.
* The far formula at `A7 = F` cannot reach `R` without crossing value blockers in columns B/C on earlier rows → excluded.

**Expected**

* `FrontierTop = ∅`
* `FrontierLeft = ∅`
* `C = ∅`
* `InRectFormulas = ∅`

---

## Notes & Implementation Hints

* You only need **presence** of non-empty cells for blocking; you do **not** need to evaluate values to know they block.
* The frontier selection guarantees at most **one** candidate per column (top) and per row (left).
* If your dependency resolver detects that any frontier candidate **does not** actually spill into `R` (e.g., size too small), that’s fine—its evaluation was necessary to know that.
* For huge sheets with open-ended edges, index `O` by:

  * `map<col, sorted rows>` and `map<row, sorted cols>`,
  * plus a quick way to enumerate rows/cols that intersect `R` (bounded) or that exist at/after `rT`/`cL` (open-ended).
* Cycles and spill errors should be handled by your existing evaluator; this spec only narrows the **set to evaluate** pre-aggregate.

If you want, I can also turn this into a short set of formal invariants and pre/post-conditions you can assert in tests.
