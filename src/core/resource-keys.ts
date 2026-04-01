export function getNamedExpressionResourceKey(opts: {
  expressionName: string;
  workbookName?: string;
  sheetName?: string;
}): string {
  if (opts.sheetName && opts.workbookName) {
    return `resource:named:sheet:${opts.workbookName}:${opts.sheetName}:${opts.expressionName}`;
  }
  if (opts.workbookName) {
    return `resource:named:workbook:${opts.workbookName}:${opts.expressionName}`;
  }
  return `resource:named:global:${opts.expressionName}`;
}

export function getTableResourceKey(opts: {
  workbookName: string;
  tableName: string;
}): string {
  return `resource:table:${opts.workbookName}:${opts.tableName}`;
}

export function getSheetResourceKey(opts: {
  workbookName: string;
  sheetName: string;
}): string {
  return `resource:sheet:${opts.workbookName}:${opts.sheetName}`;
}

export function getWorkbookResourceKey(workbookName: string): string {
  return `resource:workbook:${workbookName}`;
}
