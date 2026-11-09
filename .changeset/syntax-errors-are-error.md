---
"@ricsam/formula-engine": patch
---

A formula that does not parse shows `#ERROR!`. Some syntax errors showed
another error value because their message mentioned it, such as `#REF!` for
`=Sheet1!` ("Expected cell reference after Sheet1!").
