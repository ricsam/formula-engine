---
"@ricsam/formula-engine": minor
---

Add `WorkbookData` bulk import and change `addWorkbook` to an options object.

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
