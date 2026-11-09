---
"@ricsam/formula-engine": patch
---

Keep every table column when `setSheetContent` reorders a table's header names.
The columns were renamed one at a time, so writing `Price, Qty` over the header
row `Qty, Price` gave both columns the name `Price` midway and the table lost
one of them; structured references to the lost column read as empty. The
columns are now renamed together, and references follow their columns as with
any other header edit. This affected writing a dataset whose columns come in a
different order into a table, for example with `updateTable` and
`setSheetContent` in one `transact` call.
