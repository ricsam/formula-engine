A canonical formula is always <cell|row|col|INFINITY>, with the start written as a full cell, the end chosen according to openness, absolutes preserved on finite edges, and Excel shorthands expanded to explicit cell forms.

---

# 📐 Canonical Formula Rules

Every **range** is a **rectangle** defined by four edges:

* **Top row** (always finite, ≥ 1)
* **Left column** (always finite, ≥ A/1)
* **Bottom row** (finite or `INFINITY`)
* **Right column** (finite or `INFINITY`)

Absolutes (`$`) may be applied to any finite edge.
`INFINITY` is always implicit absolute.

---

## 1. **Start always explicit cell**

* The **start** of a range is always written as a **cell reference** (`A5`, `$A$5`, etc.).
* If the user typed a **row-only** (`5:10`) → normalize to `A5:10`.
* If the user typed a **col-only** (`A:D`) → normalize to `A1:D`.

So: *row-only → column A, col-only → row 1*.

---

## 2. **End varies by openness**

* If both bottom & right are finite → **closed rectangle**: `A5:D10`.
* If right is `INFINITY` (col-open) → **row-bounded range**: `A5:10`.
* If bottom is `INFINITY` (row-open) → **col-bounded range**: `A5:D`.
* If bottom & right are `INFINITY` → **open both**: `A5:INFINITY`.

---

## 3. **Absolute rules**

* `$` before a column letter → column is absolute.
* `$` before a row number → row is absolute.
* Applies independently at start and end.
* Example: `$A5:D$10` fixes col A and row 10, others relative.
* `INFINITY` never takes `$`.

---

## 5. **Special compatibility forms**

Excel-style shorthands (`5:10`, `A:D`, `5:5`, `A:A`, `:`) are **accepted** but **normalized** to canonical formulas by inserting implicit `A` (for rows) or `1` (for cols), and/or replacing empty with `INFINITY`.

Examples:

* `5:10` → `A5:10`
* `A:D` → `A1:D`
* `5:5` → `A5:5`
* `A:A` → `A1:A`

---


### Legend

* **Open↓** = bottom is `INFINITY` (row-open)
* **Open→** = right is `INFINITY` (col-open)
* **Start abs** = `$` on start col/row (`none`, `$col`, `$row`, `$col$row`)
* **End abs** = `$` on end col/row (when finite)
* Examples anchored at `A5` (start), and when finite end is row `10` or col `D` or cell `D10`.

---

# 🔲 Closed (Open↓=0, Open→=0) → **16 cases**

| Open↓ | Open→ | Start abs  | End abs          | Formula      |
| :---: | :---: | ---------- | ---------------- | ------------ |
|   0   |   0   | none       | none             | `A5:D10`     |
|   0   |   0   | none       | `$endCol`        | `A5:$D10`    |
|   0   |   0   | none       | `$endRow`        | `A5:D$10`    |
|   0   |   0   | none       | `$endCol$endRow` | `A5:$D$10`   |
|   0   |   0   | `$col`     | none             | `$A5:D10`    |
|   0   |   0   | `$col`     | `$endCol`        | `$A5:$D10`   |
|   0   |   0   | `$col`     | `$endRow`        | `$A5:D$10`   |
|   0   |   0   | `$col`     | `$endCol$endRow` | `$A5:$D$10`  |
|   0   |   0   | `$row`     | none             | `A$5:D10`    |
|   0   |   0   | `$row`     | `$endCol`        | `A$5:$D10`   |
|   0   |   0   | `$row`     | `$endRow`        | `A$5:D$10`   |
|   0   |   0   | `$row`     | `$endCol$endRow` | `A$5:$D$10`  |
|   0   |   0   | `$col$row` | none             | `$A$5:D10`   |
|   0   |   0   | `$col$row` | `$endCol`        | `$A$5:$D10`  |
|   0   |   0   | `$col$row` | `$endRow`        | `$A$5:D$10`  |
|   0   |   0   | `$col$row` | `$endCol$endRow` | `$A$5:$D$10` |

---

# 🔲 Open→ (Open↓=0, Open→=1) → **8 cases**

| Open↓ | Open→ | Start abs  | End abs   | Formula    |
| :---: | :---: | ---------- | --------- | ---------- |
|   0   |   1   | none       | none      | `A5:10`    |
|   0   |   1   | none       | `$endRow` | `A5:$10`   |
|   0   |   1   | `$col`     | none      | `$A5:10`   |
|   0   |   1   | `$col`     | `$endRow` | `$A5:$10`  |
|   0   |   1   | `$row`     | none      | `A$5:10`   |
|   0   |   1   | `$row`     | `$endRow` | `A$5:$10`  |
|   0   |   1   | `$col$row` | none      | `$A$5:10`  |
|   0   |   1   | `$col$row` | `$endRow` | `$A$5:$10` |

---

# 🔲 Open↓ (Open↓=1, Open→=0) → **8 cases**

| Open↓ | Open→ | Start abs  | End abs   | Formula   |
| :---: | :---: | ---------- | --------- | --------- |
|   1   |   0   | none       | none      | `A5:D`    |
|   1   |   0   | none       | `$endCol` | `A5:$D`   |
|   1   |   0   | `$col`     | none      | `$A5:D`   |
|   1   |   0   | `$col`     | `$endCol` | `$A5:$D`  |
|   1   |   0   | `$row`     | none      | `A$5:D`   |
|   1   |   0   | `$row`     | `$endCol` | `A$5:$D`  |
|   1   |   0   | `$col$row` | none      | `$A$5:D`  |
|   1   |   0   | `$col$row` | `$endCol` | `$A$5:$D` |

---

# 🔲 Open both (Open↓=1, Open→=1) → **4 cases**

| Open↓ | Open→ | Start abs  | End abs | Formula         |
| :---: | :---: | ---------- | ------- | --------------- |
|   1   |   1   | none       | (n/a)   | `A5:INFINITY`   |
|   1   |   1   | `$col`     | (n/a)   | `$A5:INFINITY`  |
|   1   |   1   | `$row`     | (n/a)   | `A$5:INFINITY`  |
|   1   |   1   | `$col$row` | (n/a)   | `$A$5:INFINITY` |

---

## ✅ Totals

* Closed: 16
* Open→: 8
* Open↓: 8
* Open both: 4
  **= 36 combinations**


**Excel-compatibility mapping** (including partial absolutes) to your canonical forms:

# Whole row(s)

| Input    | Meaning                   | Canonical form |
| -------- | ------------------------- | -------------- |
| `5:5`    | Entire row 5              | `A5:5`         |
| `$5:5`   | Row 5, absolute start row | `A$5:5`        |
| `5:$5`   | Row 5, absolute end row   | `A5:$5`        |
| `$5:$5`  | Row 5, both ends absolute | `A$5:$5`       |
| `5:10`   | Rows 5–10 (all columns)   | `A5:10`        |
| `$5:10`  | Abs start row             | `A$5:10`       |
| `5:$10`  | Abs end row               | `A5:$10`       |
| `$5:$10` | Abs both rows             | `A$5:$10`      |

# Whole column(s)

| Input   | Meaning                | Canonical form |
| ------- | ---------------------- | -------------- |
| `A:A`   | Entire column A        | `A1:A`         |
| `$A:A`  | Abs start col          | `$A1:A`        |
| `A:$A`  | Abs end col            | `A1:$A`        |
| `$A:$A` | Abs both cols          | `$A1:$A`       |
| `A:D`   | Columns A–D (all rows) | `A1:D`         |
| `$A:D`  | Abs start col          | `$A1:D`        |
| `A:$D`  | Abs end col            | `A1:$D`        |
| `$A:$D` | Abs both cols          | `$A1:$D`       |
