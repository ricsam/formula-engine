import type { StructuredReferenceNode } from "../parser/ast";
import { functions } from "../functions/function-registry";
import { columnToIndex } from "../core/utils";
import {
  FormulaError,
  type CellAddress,
  type RangeAddress,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "../core/types";
import { Lexer, type Token, type TokenType } from "../parser/lexer";
import { ParseError, parseFormula } from "../parser/parser";

/** A half-open UTF-16 range in the exact formula string supplied by the caller. */
export interface SourceSpan {
  start: number;
  end: number;
}

export type FormulaTokenKind =
  | "formula-marker"
  | "function"
  | "number"
  | "string"
  | "boolean"
  | "error"
  | "operator"
  | "delimiter"
  | "workbook-name"
  | "sheet-name"
  | "cell-reference"
  | "table-name"
  | "table-column"
  | "table-selector"
  | "named-expression"
  | "keyword"
  | "invalid";

export type FormulaTokenModifier =
  | "absolute"
  | "unknown"
  | "external"
  | "current-row";

/** An editor-neutral semantic token. Tokens never overlap. */
export interface FormulaToken {
  span: SourceSpan;
  kind: FormulaTokenKind;
  modifiers: readonly FormulaTokenModifier[];
  /** Correlates formula text with an entry in FormulaAnalysis.references. */
  referenceId?: string;
}

export type NamedExpressionScope =
  | { type: "global" }
  | { type: "workbook"; workbookName: string }
  | { type: "sheet"; workbookName: string; sheetName: string };

export type FormulaReferenceTarget =
  | { type: "cell"; address: CellAddress }
  | { type: "range"; address: RangeAddress }
  | {
      type: "table";
      workbookName: string;
      sheetName: string;
      tableName: string;
      range: SpreadsheetRange;
    }
  | {
      type: "named-expression";
      name: string;
      scope: NamedExpressionScope;
    };

export type FormulaReferenceResolution =
  | {
      status: "resolved";
      targets: readonly FormulaReferenceTarget[];
    }
  | {
      status: "unresolved";
      reason:
        | "missing-origin"
        | "missing-workbook"
        | "missing-sheet"
        | "missing-table"
        | "missing-column"
        | "missing-name";
    }
  | { status: "dynamic" };

export interface FormulaReference {
  /** Stable within one analysis result. */
  id: string;
  span: SourceSpan;
  kind:
    | "cell"
    | "range"
    | "structured-reference"
    | "named-expression"
    | "3d-range";
  resolution: FormulaReferenceResolution;
}

export interface FormulaDiagnostic {
  span: SourceSpan;
  severity: "error" | "warning";
  code: string;
  message: string;
}

export interface FormulaAnalysis {
  /** The exact input string. All spans index this string. */
  formula: string;
  tokens: readonly FormulaToken[];
  references: readonly FormulaReference[];
  diagnostics: readonly FormulaDiagnostic[];
}

export interface FormulaAnalysisOptions {
  /** A formula with or without a leading equals sign. */
  formula: string;
  /** The cell from which relative scopes and structured references resolve. */
  origin?: CellAddress;
}

/** @internal Resolution hooks supplied by FormulaEngine. */
interface FormulaAnalysisEnvironment {
  hasWorkbook(workbookName: string): boolean;
  hasSheet(workbookName: string, sheetName: string): boolean;
  getOrderedSheetNames(workbookName: string): readonly string[];
  getTable(
    workbookName: string,
    tableName: string
  ): TableDefinition | undefined;
  findTableForCell(address: CellAddress): TableDefinition | undefined;
  resolveNamedExpression(options: {
    name: string;
    workbookName?: string;
    sheetName?: string;
    origin?: CellAddress;
  }): NamedExpressionScope | undefined;
}

interface WorkingToken {
  source: Token;
  span: SourceSpan;
  kind: FormulaTokenKind;
  modifiers: Set<FormulaTokenModifier>;
  referenceId?: string;
}

type CellEndpoint = {
  startIndex: number;
  endIndex: number;
  span: SourceSpan;
  raw: string;
  type: "cell" | "column" | "row" | "infinity";
  colIndex?: number;
  rowIndex?: number;
  absolute: boolean;
};

type ReferenceQualifier = {
  startIndex: number;
  endIndex: number;
  endpointIndex: number;
  workbookName?: string;
  sheetName?: string;
  startSheetName?: string;
  endSheetName?: string;
  workbookTokenIndexes: number[];
  sheetTokenIndexes: number[];
};

const TOKEN_KIND: Record<TokenType, FormulaTokenKind> = {
  NUMBER: "number",
  STRING: "string",
  BOOLEAN: "boolean",
  IDENTIFIER: "named-expression",
  FUNCTION: "function",
  OPERATOR: "operator",
  LPAREN: "delimiter",
  RPAREN: "delimiter",
  LBRACE: "delimiter",
  RBRACE: "delimiter",
  LBRACKET: "delimiter",
  RBRACKET: "delimiter",
  COMMA: "delimiter",
  SEMICOLON: "delimiter",
  COLON: "delimiter",
  DOLLAR: "operator",
  EXCLAMATION: "delimiter",
  AT: "keyword",
  HASH: "table-selector",
  INFINITY: "keyword",
  ERROR: "error",
  EOF: "invalid",
  WHITESPACE: "delimiter",
};

const KNOWN_FORMULA_ERRORS = new Set<string>(Object.values(FormulaError));
const ENDPOINT_TOKEN_TYPES = new Set<TokenType>([
  "DOLLAR",
  "IDENTIFIER",
  "NUMBER",
  "INFINITY",
]);
const POSITIVE_INFINITY: SpreadsheetRangeEnd = {
  type: "infinity",
  sign: "positive",
};

function normalizeQuotedName(value: string): string {
  if (value.startsWith("'") && value.endsWith("'")) {
    return value.slice(1, -1).replace(/''/g, "'");
  }
  return value;
}

function isContiguous(left: WorkingToken, right: WorkingToken): boolean {
  return left.span.end === right.span.start;
}

function addModifier(
  tokens: WorkingToken[],
  indexes: readonly number[],
  modifier: FormulaTokenModifier
): void {
  for (const index of indexes) {
    tokens[index]?.modifiers.add(modifier);
  }
}

function markTokens(
  tokens: WorkingToken[],
  startIndex: number,
  endIndex: number,
  kind: FormulaTokenKind,
  referenceId?: string
): void {
  for (let index = startIndex; index < endIndex; index++) {
    const token = tokens[index];
    if (!token) continue;
    token.kind = kind;
    if (referenceId) token.referenceId = referenceId;
  }
}

function assignReferenceId(
  tokens: WorkingToken[],
  startIndex: number,
  endIndex: number,
  referenceId: string
): void {
  for (let index = startIndex; index < endIndex; index++) {
    const token = tokens[index];
    if (token) token.referenceId = referenceId;
  }
}

function findMatchingBracket(
  tokens: WorkingToken[],
  startIndex: number
): number {
  if (tokens[startIndex]?.source.type !== "LBRACKET") return -1;
  let depth = 0;
  for (let index = startIndex; index < tokens.length; index++) {
    const type = tokens[index]?.source.type;
    if (type === "LBRACKET") depth++;
    if (type === "RBRACKET") {
      depth--;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function parseEndpoint(
  tokens: WorkingToken[],
  startIndex: number,
  occupied: ReadonlySet<number>
): CellEndpoint | undefined {
  const first = tokens[startIndex];
  if (!first || occupied.has(startIndex)) return undefined;

  let raw = "";
  let best: CellEndpoint | undefined;
  let previous: WorkingToken | undefined;

  for (
    let index = startIndex;
    index < Math.min(tokens.length, startIndex + 4);
    index++
  ) {
    const token = tokens[index];
    if (
      !token ||
      occupied.has(index) ||
      !ENDPOINT_TOKEN_TYPES.has(token.source.type) ||
      (previous && !isContiguous(previous, token))
    ) {
      break;
    }
    raw += token.source.value;
    previous = token;

    const cell = raw.match(/^(\$)?([A-Z]+)(\$)?([1-9][0-9]*)$/i);
    const column = raw.match(/^(\$)?([A-Z]+)$/i);
    const row = raw.match(/^(\$)?([1-9][0-9]*)$/);
    const infinity = /^INFINITY$/i.test(raw);
    const base = {
      startIndex,
      endIndex: index + 1,
      span: { start: first.span.start, end: token.span.end },
      raw,
      absolute: raw.includes("$"),
    };

    if (cell?.[2] && cell[4]) {
      best = {
        ...base,
        type: "cell",
        colIndex: columnToIndex(cell[2]),
        rowIndex: Number(cell[4]) - 1,
      };
    } else if (column?.[2]) {
      best = {
        ...base,
        type: "column",
        colIndex: columnToIndex(column[2]),
      };
    } else if (row?.[2]) {
      best = {
        ...base,
        type: "row",
        rowIndex: Number(row[2]) - 1,
      };
    } else if (infinity) {
      best = { ...base, type: "infinity" };
    }
  }

  return best;
}

function rangeFromEndpoints(
  start: CellEndpoint,
  end: CellEndpoint
): SpreadsheetRange | undefined {
  if (start.type === "cell") {
    const rangeStart = { col: start.colIndex!, row: start.rowIndex! };
    if (end.type === "cell") {
      return {
        start: {
          col: Math.min(rangeStart.col, end.colIndex!),
          row: Math.min(rangeStart.row, end.rowIndex!),
        },
        end: {
          col: {
            type: "number",
            value: Math.max(rangeStart.col, end.colIndex!),
          },
          row: {
            type: "number",
            value: Math.max(rangeStart.row, end.rowIndex!),
          },
        },
      };
    }
    if (end.type === "column") {
      return {
        start: {
          col: Math.min(rangeStart.col, end.colIndex!),
          row: rangeStart.row,
        },
        end: {
          col: {
            type: "number",
            value: Math.max(rangeStart.col, end.colIndex!),
          },
          row: POSITIVE_INFINITY,
        },
      };
    }
    if (end.type === "row") {
      return {
        start: {
          col: rangeStart.col,
          row: Math.min(rangeStart.row, end.rowIndex!),
        },
        end: {
          col: POSITIVE_INFINITY,
          row: {
            type: "number",
            value: Math.max(rangeStart.row, end.rowIndex!),
          },
        },
      };
    }
    if (end.type === "infinity") {
      return {
        start: rangeStart,
        end: { col: POSITIVE_INFINITY, row: POSITIVE_INFINITY },
      };
    }
  }

  if (start.type === "column" && end.type === "column") {
    return {
      start: { col: Math.min(start.colIndex!, end.colIndex!), row: 0 },
      end: {
        col: {
          type: "number",
          value: Math.max(start.colIndex!, end.colIndex!),
        },
        row: POSITIVE_INFINITY,
      },
    };
  }

  if (start.type === "row" && end.type === "row") {
    return {
      start: { col: 0, row: Math.min(start.rowIndex!, end.rowIndex!) },
      end: {
        col: POSITIVE_INFINITY,
        row: {
          type: "number",
          value: Math.max(start.rowIndex!, end.rowIndex!),
        },
      },
    };
  }

  return undefined;
}

function parseQualifier(
  formula: string,
  tokens: WorkingToken[],
  startIndex: number,
  occupied: ReadonlySet<number>
): ReferenceQualifier | undefined {
  const first = tokens[startIndex];
  if (!first || occupied.has(startIndex)) return undefined;

  if (first.source.type === "LBRACKET") {
    const closeIndex = findMatchingBracket(tokens, startIndex);
    if (closeIndex < 0) return undefined;
    const close = tokens[closeIndex]!;
    const sheet = tokens[closeIndex + 1];
    const separator = tokens[closeIndex + 2];
    if (sheet?.source.type !== "IDENTIFIER") return undefined;

    if (separator?.source.type === "EXCLAMATION") {
      return {
        startIndex,
        endIndex: closeIndex + 3,
        endpointIndex: closeIndex + 3,
        workbookName: formula.slice(first.span.end, close.span.start),
        sheetName: normalizeQuotedName(sheet.source.value),
        workbookTokenIndexes: Array.from(
          { length: Math.max(0, closeIndex - startIndex - 1) },
          (_, index) => startIndex + index + 1
        ),
        sheetTokenIndexes: [closeIndex + 1],
      };
    }

    const colon = tokens[closeIndex + 2];
    const endSheet = tokens[closeIndex + 3];
    const exclamation = tokens[closeIndex + 4];
    if (
      colon?.source.type === "COLON" &&
      endSheet?.source.type === "IDENTIFIER" &&
      exclamation?.source.type === "EXCLAMATION"
    ) {
      return {
        startIndex,
        endIndex: closeIndex + 5,
        endpointIndex: closeIndex + 5,
        workbookName: formula.slice(first.span.end, close.span.start),
        startSheetName: normalizeQuotedName(sheet.source.value),
        endSheetName: normalizeQuotedName(endSheet.source.value),
        workbookTokenIndexes: Array.from(
          { length: Math.max(0, closeIndex - startIndex - 1) },
          (_, index) => startIndex + index + 1
        ),
        sheetTokenIndexes: [closeIndex + 1, closeIndex + 3],
      };
    }
    return undefined;
  }

  if (first.source.type !== "IDENTIFIER") return undefined;
  const second = tokens[startIndex + 1];
  if (second?.source.type === "EXCLAMATION") {
    return {
      startIndex,
      endIndex: startIndex + 2,
      endpointIndex: startIndex + 2,
      sheetName: normalizeQuotedName(first.source.value),
      workbookTokenIndexes: [],
      sheetTokenIndexes: [startIndex],
    };
  }

  const endSheet = tokens[startIndex + 2];
  const exclamation = tokens[startIndex + 3];
  if (
    second?.source.type === "COLON" &&
    endSheet?.source.type === "IDENTIFIER" &&
    exclamation?.source.type === "EXCLAMATION"
  ) {
    return {
      startIndex,
      endIndex: startIndex + 4,
      endpointIndex: startIndex + 4,
      startSheetName: normalizeQuotedName(first.source.value),
      endSheetName: normalizeQuotedName(endSheet.source.value),
      workbookTokenIndexes: [],
      sheetTokenIndexes: [startIndex, startIndex + 2],
    };
  }

  return undefined;
}

function unresolved(
  reason: Extract<
    FormulaReferenceResolution,
    { status: "unresolved" }
  >["reason"]
): FormulaReferenceResolution {
  return { status: "unresolved", reason };
}

function validateScope(
  workbookName: string | undefined,
  sheetName: string | undefined,
  environment: FormulaAnalysisEnvironment | undefined
): FormulaReferenceResolution | undefined {
  if (!workbookName || !sheetName) return unresolved("missing-origin");
  if (environment && !environment.hasWorkbook(workbookName)) {
    return unresolved("missing-workbook");
  }
  if (environment && !environment.hasSheet(workbookName, sheetName)) {
    return unresolved("missing-sheet");
  }
  return undefined;
}

function resolveCellReference(
  start: CellEndpoint,
  end: CellEndpoint | undefined,
  qualifier: ReferenceQualifier | undefined,
  origin: CellAddress | undefined,
  environment: FormulaAnalysisEnvironment | undefined
): FormulaReferenceResolution {
  const workbookName = qualifier?.workbookName ?? origin?.workbookName;
  const is3d = Boolean(qualifier?.startSheetName && qualifier.endSheetName);

  if (is3d) {
    if (!workbookName) return unresolved("missing-origin");
    if (environment && !environment.hasWorkbook(workbookName)) {
      return unresolved("missing-workbook");
    }
    if (!environment) return unresolved("missing-sheet");
    const names = environment.getOrderedSheetNames(workbookName);
    const startIndex = names.indexOf(qualifier!.startSheetName!);
    const endIndex = names.indexOf(qualifier!.endSheetName!);
    if (startIndex < 0 || endIndex < 0) return unresolved("missing-sheet");
    const low = Math.min(startIndex, endIndex);
    const high = Math.max(startIndex, endIndex);
    const range = end ? rangeFromEndpoints(start, end) : undefined;
    const targets: FormulaReferenceTarget[] = [];
    for (const sheetName of names.slice(low, high + 1)) {
      if (range) {
        targets.push({
          type: "range",
          address: { workbookName, sheetName, range },
        });
      } else {
        targets.push({
          type: "cell",
          address: {
            workbookName,
            sheetName,
            colIndex: start.colIndex!,
            rowIndex: start.rowIndex!,
          },
        });
      }
    }
    return { status: "resolved", targets };
  }

  const sheetName = qualifier?.sheetName ?? origin?.sheetName;
  const invalidScope = validateScope(workbookName, sheetName, environment);
  if (invalidScope) return invalidScope;

  if (end) {
    const range = rangeFromEndpoints(start, end);
    if (!range) return unresolved("missing-origin");
    return {
      status: "resolved",
      targets: [
        {
          type: "range",
          address: {
            workbookName: workbookName!,
            sheetName: sheetName!,
            range,
          },
        },
      ],
    };
  }

  return {
    status: "resolved",
    targets: [
      {
        type: "cell",
        address: {
          workbookName: workbookName!,
          sheetName: sheetName!,
          colIndex: start.colIndex!,
          rowIndex: start.rowIndex!,
        },
      },
    ],
  };
}

function tableRange(
  table: TableDefinition,
  node: StructuredReferenceNode,
  origin: CellAddress | undefined
): SpreadsheetRange | undefined {
  let startRow = table.start.rowIndex + 1;
  let endRow = table.endRow;
  if (node.selector === "#Headers") {
    startRow = table.start.rowIndex;
    endRow = { type: "number", value: table.start.rowIndex };
  } else if (node.selector === "#All") {
    startRow = table.start.rowIndex;
  } else if (node.selector === "#Data") {
    startRow = table.start.rowIndex + 1;
  } else if (node.isCurrentRow) {
    if (!origin) return undefined;
    startRow = origin.rowIndex;
    endRow = { type: "number", value: origin.rowIndex };
  }

  let startCol = table.start.colIndex;
  let endCol = table.start.colIndex + table.headers.size - 1;
  if (node.cols) {
    const startHeader = table.headers.get(node.cols.startCol);
    const endHeader = table.headers.get(node.cols.endCol);
    if (!startHeader || !endHeader) return undefined;
    startCol = table.start.colIndex + startHeader.index;
    endCol = table.start.colIndex + endHeader.index;
  }

  return {
    start: { col: startCol, row: startRow },
    end: { col: { type: "number", value: endCol }, row: endRow },
  };
}

function resolveStructuredReference(
  node: StructuredReferenceNode,
  origin: CellAddress | undefined,
  environment: FormulaAnalysisEnvironment | undefined
): FormulaReferenceResolution {
  const workbookName = node.workbookName ?? origin?.workbookName;
  if (!workbookName) return unresolved("missing-origin");
  if (environment && !environment.hasWorkbook(workbookName)) {
    return unresolved("missing-workbook");
  }
  if (!environment) return unresolved("missing-table");

  const table = node.tableName
    ? environment.getTable(workbookName, node.tableName)
    : origin
    ? environment.findTableForCell(origin)
    : undefined;
  if (!table) {
    return unresolved(
      node.tableName || origin ? "missing-table" : "missing-origin"
    );
  }
  if (node.isCurrentRow && !origin) return unresolved("missing-origin");

  if (node.cols) {
    if (
      !table.headers.has(node.cols.startCol) ||
      !table.headers.has(node.cols.endCol)
    ) {
      return unresolved("missing-column");
    }
  }
  const range = tableRange(table, node, origin);
  if (!range) return unresolved("missing-column");
  return {
    status: "resolved",
    targets: [
      {
        type: "table",
        workbookName: table.workbookName,
        sheetName: table.sheetName,
        tableName: table.name,
        range,
      },
    ],
  };
}

function resolveNamedReference(
  name: string,
  qualifier: ReferenceQualifier | undefined,
  origin: CellAddress | undefined,
  environment: FormulaAnalysisEnvironment | undefined
): FormulaReferenceResolution {
  if (!environment) return unresolved("missing-name");
  const workbookName = qualifier?.workbookName ?? origin?.workbookName;
  if (qualifier && !workbookName) return unresolved("missing-origin");
  if (workbookName && !environment.hasWorkbook(workbookName)) {
    return unresolved("missing-workbook");
  }
  if (
    qualifier?.sheetName &&
    workbookName &&
    !environment.hasSheet(workbookName, qualifier.sheetName)
  ) {
    return unresolved("missing-sheet");
  }
  const scope = environment.resolveNamedExpression({
    name,
    workbookName: qualifier?.workbookName,
    sheetName: qualifier?.sheetName,
    origin,
  });
  return scope
    ? {
        status: "resolved",
        targets: [{ type: "named-expression", name, scope }],
      }
    : unresolved("missing-name");
}

function diagnosticForResolution(
  reference: FormulaReference
): FormulaDiagnostic | undefined {
  if (reference.resolution.status !== "unresolved") return undefined;
  const messages: Record<
    Extract<FormulaReferenceResolution, { status: "unresolved" }>["reason"],
    string
  > = {
    "missing-origin": "An origin cell is required to resolve this reference.",
    "missing-workbook": "The referenced workbook does not exist.",
    "missing-sheet": "The referenced sheet does not exist.",
    "missing-table":
      "The referenced table does not exist or cannot be inferred.",
    "missing-column": "The referenced table column does not exist.",
    "missing-name":
      "The named expression does not exist in the applicable scope.",
  };
  return {
    span: reference.span,
    severity: "warning",
    code: `reference.${reference.resolution.reason}`,
    message: messages[reference.resolution.reason],
  };
}

function createParserDiagnostics(
  formula: string,
  body: string,
  bodyOffset: number,
  rawTokens: readonly Token[]
): FormulaDiagnostic[] {
  const diagnostics: FormulaDiagnostic[] = [];
  if (body.length === 0) {
    if (bodyOffset > 0) {
      diagnostics.push({
        span: { start: bodyOffset, end: bodyOffset },
        severity: "error",
        code: "syntax.empty-formula",
        message: "Expected an expression after '='.",
      });
    }
    return diagnostics;
  }

  for (const token of rawTokens) {
    if (token.type === "STRING") {
      const text = body.slice(token.position.start, token.position.end);
      if (!text.endsWith('"')) {
        diagnostics.push({
          span: {
            start: token.position.start + bodyOffset,
            end: token.position.end + bodyOffset,
          },
          severity: "error",
          code: "syntax.unterminated-string",
          message: "Unterminated string literal.",
        });
      }
    }
  }

  try {
    parseFormula(body);
  } catch (error) {
    const parseError = error instanceof ParseError ? error : undefined;
    const rawSpan = parseError?.position ?? { start: 0, end: body.length };
    let start = Math.max(
      0,
      Math.min(formula.length, rawSpan.start + bodyOffset)
    );
    let end = Math.max(
      start,
      Math.min(formula.length, rawSpan.end + bodyOffset)
    );
    if (start === end && formula.length > 0) {
      start = Math.max(bodyOffset, start - 1);
      end = Math.min(formula.length, Math.max(start + 1, end));
    }
    diagnostics.push({
      span: { start, end },
      severity: "error",
      code: "syntax.parse-error",
      message: error instanceof Error ? error.message : String(error),
    });
  }

  return diagnostics;
}

function structuredCandidateEnd(
  tokens: WorkingToken[],
  startIndex: number,
  occupied: ReadonlySet<number>
): number {
  if (occupied.has(startIndex)) return -1;
  const token = tokens[startIndex];
  if (!token) return -1;

  if (token.source.type === "AT") {
    const next = tokens[startIndex + 1];
    return next?.source.type === "IDENTIFIER" ? startIndex + 1 : -1;
  }

  if (token.source.type === "IDENTIFIER") {
    if (tokens[startIndex + 1]?.source.type === "LBRACKET") {
      return findMatchingBracket(tokens, startIndex + 1);
    }
    if (
      tokens[startIndex + 1]?.source.type === "EXCLAMATION" &&
      tokens[startIndex + 2]?.source.type === "IDENTIFIER" &&
      tokens[startIndex + 3]?.source.type === "LBRACKET"
    ) {
      return findMatchingBracket(tokens, startIndex + 3);
    }
    return -1;
  }

  if (token.source.type === "LBRACKET") {
    const close = findMatchingBracket(tokens, startIndex);
    if (close < 0) return -1;
    if (
      tokens[close + 1]?.source.type === "EXCLAMATION" &&
      tokens[close + 2]?.source.type === "IDENTIFIER" &&
      tokens[close + 3]?.source.type === "LBRACKET"
    ) {
      return findMatchingBracket(tokens, close + 3);
    }
    // [Book]Sheet!A1 is a workbook qualifier, not a bare table column.
    if (
      tokens[close + 1]?.source.type === "IDENTIFIER" &&
      (tokens[close + 2]?.source.type === "EXCLAMATION" ||
        tokens[close + 2]?.source.type === "COLON")
    ) {
      return -1;
    }
    return close;
  }

  return -1;
}

function markStructuredSyntax(
  tokens: WorkingToken[],
  startIndex: number,
  endIndex: number,
  node: StructuredReferenceNode,
  referenceId: string
): void {
  assignReferenceId(tokens, startIndex, endIndex + 1, referenceId);
  if (node.workbookName) {
    const firstClose = tokens.findIndex(
      (token, index) =>
        index >= startIndex &&
        index <= endIndex &&
        token.source.type === "RBRACKET"
    );
    if (firstClose > startIndex) {
      markTokens(
        tokens,
        startIndex + 1,
        firstClose,
        "workbook-name",
        referenceId
      );
      addModifier(
        tokens,
        Array.from(
          { length: firstClose - startIndex - 1 },
          (_, i) => startIndex + i + 1
        ),
        "external"
      );
    }
  }

  let tableNameMarked = false;
  for (let index = startIndex; index <= endIndex; index++) {
    const token = tokens[index]!;
    if (
      node.tableName &&
      !tableNameMarked &&
      token.source.type === "IDENTIFIER" &&
      token.source.value === node.tableName &&
      tokens[index + 1]?.source.type === "LBRACKET"
    ) {
      token.kind = "table-name";
      tableNameMarked = true;
      continue;
    }
    if (
      node.sheetName &&
      token.source.type === "IDENTIFIER" &&
      normalizeQuotedName(token.source.value) === node.sheetName &&
      tokens[index + 1]?.source.type === "EXCLAMATION"
    ) {
      token.kind = "sheet-name";
      continue;
    }
    if (token.source.type === "HASH") {
      token.kind = "table-selector";
      const selectorName = tokens[index + 1];
      if (selectorName?.source.type === "IDENTIFIER") {
        selectorName.kind = "table-selector";
      }
      continue;
    }
    if (token.source.type === "AT") {
      token.kind = "keyword";
      token.modifiers.add("current-row");
      continue;
    }
    if (
      token.source.type !== "LBRACKET" &&
      token.source.type !== "RBRACKET" &&
      token.source.type !== "COMMA" &&
      token.source.type !== "COLON" &&
      token.source.type !== "EXCLAMATION" &&
      token.kind !== "workbook-name" &&
      token.kind !== "sheet-name" &&
      token.kind !== "table-name" &&
      token.kind !== "table-selector"
    ) {
      token.kind = "table-column";
      if (node.isCurrentRow) token.modifiers.add("current-row");
    }
  }
}

/**
 * Analyze formula text without requiring an engine. Static cell references can
 * resolve when origin is supplied; engine-owned resources remain unresolved.
 */
export function analyzeFormula(
  options: FormulaAnalysisOptions
): FormulaAnalysis {
  return analyzeFormulaWithEnvironment(options);
}

/** @internal Used by FormulaEngine to add workbook-aware resolution. */
export function analyzeFormulaWithEnvironment(
  options: FormulaAnalysisOptions,
  environment?: FormulaAnalysisEnvironment
): FormulaAnalysis {
  const formula = options.formula;
  try {
    const bodyOffset = formula.startsWith("=") ? 1 : 0;
    const body = formula.slice(bodyOffset);
    const rawTokens = new Lexer(body)
      .tokenize()
      .filter((token) => token.type !== "EOF" && token.type !== "WHITESPACE");
    const tokens: WorkingToken[] = rawTokens.map((source) => ({
      source,
      span: {
        start: source.position.start + bodyOffset,
        end: source.position.end + bodyOffset,
      },
      kind:
        source.type === "ERROR" && !KNOWN_FORMULA_ERRORS.has(source.value)
          ? "invalid"
          : TOKEN_KIND[source.type],
      modifiers: new Set<FormulaTokenModifier>(),
    }));
    const occupied = new Set<number>();
    const references: FormulaReference[] = [];
    const diagnostics = createParserDiagnostics(
      formula,
      body,
      bodyOffset,
      rawTokens
    );

    for (let index = 0; index < tokens.length; index++) {
      const token = tokens[index]!;
      if (token.source.type === "FUNCTION") {
        if (!functions[token.source.value.toUpperCase()]) {
          token.modifiers.add("unknown");
          diagnostics.push({
            span: token.span,
            severity: "warning",
            code: "function.unknown",
            message: `Unknown function ${token.source.value}.`,
          });
        }
      }
      if (token.kind === "invalid") {
        diagnostics.push({
          span: token.span,
          severity: "error",
          code: "syntax.invalid-token",
          message: `Invalid token ${JSON.stringify(
            formula.slice(token.span.start, token.span.end)
          )}.`,
        });
      }
    }

    // Structured references are scanned first because their identifiers can
    // otherwise look like ordinary A1 or named-expression references.
    for (let index = 0; index < tokens.length; index++) {
      if (occupied.has(index)) continue;
      const endIndex = structuredCandidateEnd(tokens, index, occupied);
      if (endIndex < index) continue;
      const startToken = tokens[index]!;
      const endToken = tokens[endIndex]!;
      const candidate = formula.slice(startToken.span.start, endToken.span.end);
      let node: StructuredReferenceNode | undefined;
      try {
        const ast = parseFormula(candidate);
        if (ast.type === "structured-reference") node = ast;
      } catch {
        // A partially typed structured reference is still represented by its
        // lexical tokens and the parser diagnostic above.
      }
      if (!node) continue;

      const id = `reference:${references.length}`;
      const resolution = resolveStructuredReference(
        node,
        options.origin,
        environment
      );
      const reference: FormulaReference = {
        id,
        span: { start: startToken.span.start, end: endToken.span.end },
        kind: "structured-reference",
        resolution,
      };
      references.push(reference);
      for (let cursor = index; cursor <= endIndex; cursor++)
        occupied.add(cursor);
      markStructuredSyntax(tokens, index, endIndex, node, id);
      if (resolution.status === "unresolved") {
        addModifier(
          tokens,
          Array.from({ length: endIndex - index + 1 }, (_, i) => index + i),
          "unknown"
        );
      }
      const diagnostic = diagnosticForResolution(reference);
      if (diagnostic) diagnostics.push(diagnostic);
      index = endIndex;
    }

    for (let index = 0; index < tokens.length; index++) {
      if (occupied.has(index)) continue;
      const qualifier = parseQualifier(formula, tokens, index, occupied);
      const endpointIndex = qualifier?.endpointIndex ?? index;
      const endpoint = parseEndpoint(tokens, endpointIndex, occupied);

      if (qualifier) {
        markTokens(
          tokens,
          qualifier.startIndex,
          qualifier.endIndex,
          "delimiter"
        );
        markTokensForQualifier(tokens, qualifier);
      }

      if (endpoint) {
        let rangeEnd: CellEndpoint | undefined;
        const colonIndex = endpoint.endIndex;
        if (tokens[colonIndex]?.source.type === "COLON") {
          const possibleEnd = parseEndpoint(tokens, colonIndex + 1, occupied);
          if (possibleEnd && rangeFromEndpoints(endpoint, possibleEnd)) {
            rangeEnd = possibleEnd;
          }
        }

        const validSingleCell = endpoint.type === "cell";
        if (validSingleCell || rangeEnd) {
          const referenceEndIndex = rangeEnd?.endIndex ?? endpoint.endIndex;
          const id = `reference:${references.length}`;
          const resolution = resolveCellReference(
            endpoint,
            rangeEnd,
            qualifier,
            options.origin,
            environment
          );
          const startSpan = qualifier
            ? tokens[qualifier.startIndex]!.span.start
            : endpoint.span.start;
          const endSpan = (rangeEnd ?? endpoint).span.end;
          const reference: FormulaReference = {
            id,
            span: { start: startSpan, end: endSpan },
            kind: qualifier?.startSheetName
              ? "3d-range"
              : rangeEnd
              ? "range"
              : "cell",
            resolution,
          };
          references.push(reference);
          const occupiedStart = qualifier?.startIndex ?? endpoint.startIndex;
          for (
            let cursor = occupiedStart;
            cursor < referenceEndIndex;
            cursor++
          ) {
            occupied.add(cursor);
          }
          assignReferenceId(tokens, occupiedStart, referenceEndIndex, id);
          markTokens(
            tokens,
            endpoint.startIndex,
            endpoint.endIndex,
            "cell-reference",
            id
          );
          if (endpoint.absolute) {
            addModifier(
              tokens,
              Array.from(
                { length: endpoint.endIndex - endpoint.startIndex },
                (_, i) => endpoint.startIndex + i
              ),
              "absolute"
            );
          }
          if (rangeEnd) {
            markTokens(
              tokens,
              rangeEnd.startIndex,
              rangeEnd.endIndex,
              "cell-reference",
              id
            );
            if (rangeEnd.absolute) {
              addModifier(
                tokens,
                Array.from(
                  { length: rangeEnd.endIndex - rangeEnd.startIndex },
                  (_, i) => rangeEnd.startIndex + i
                ),
                "absolute"
              );
            }
          }
          if (qualifier?.workbookName) {
            addModifier(
              tokens,
              Array.from(
                { length: referenceEndIndex - occupiedStart },
                (_, i) => occupiedStart + i
              ),
              "external"
            );
          }
          if (resolution.status === "unresolved") {
            addModifier(
              tokens,
              Array.from(
                { length: referenceEndIndex - occupiedStart },
                (_, i) => occupiedStart + i
              ),
              "unknown"
            );
          }
          const diagnostic = diagnosticForResolution(reference);
          if (diagnostic) diagnostics.push(diagnostic);
          index = referenceEndIndex - 1;
          continue;
        }
      }

      const nameToken = tokens[endpointIndex];
      if (nameToken?.source.type === "IDENTIFIER") {
        const id = `reference:${references.length}`;
        const resolution = resolveNamedReference(
          nameToken.source.value,
          qualifier,
          options.origin,
          environment
        );
        const occupiedStart = qualifier?.startIndex ?? endpointIndex;
        const endIndex = endpointIndex + 1;
        const reference: FormulaReference = {
          id,
          span: {
            start: tokens[occupiedStart]!.span.start,
            end: nameToken.span.end,
          },
          kind: "named-expression",
          resolution,
        };
        references.push(reference);
        for (let cursor = occupiedStart; cursor < endIndex; cursor++) {
          occupied.add(cursor);
        }
        assignReferenceId(tokens, occupiedStart, endIndex, id);
        nameToken.kind = "named-expression";
        if (qualifier?.workbookName) {
          addModifier(
            tokens,
            Array.from(
              { length: endIndex - occupiedStart },
              (_, i) => occupiedStart + i
            ),
            "external"
          );
        }
        if (resolution.status === "unresolved") {
          addModifier(
            tokens,
            Array.from(
              { length: endIndex - occupiedStart },
              (_, i) => occupiedStart + i
            ),
            "unknown"
          );
        }
        const diagnostic = diagnosticForResolution(reference);
        if (diagnostic) diagnostics.push(diagnostic);
        index = endIndex - 1;
      }
    }

    const publicTokens: FormulaToken[] = tokens.map((token) => ({
      span: token.span,
      kind: token.kind,
      modifiers: Array.from(token.modifiers).sort(),
      ...(token.referenceId ? { referenceId: token.referenceId } : {}),
    }));
    if (bodyOffset > 0) {
      publicTokens.unshift({
        span: { start: 0, end: 1 },
        kind: "formula-marker",
        modifiers: [],
      });
    }

    diagnostics.sort(
      (left, right) =>
        left.span.start - right.span.start ||
        left.span.end - right.span.end ||
        left.code.localeCompare(right.code)
    );
    return { formula, tokens: publicTokens, references, diagnostics };
  } catch (error) {
    // Formula analysis runs on every keystroke and is deliberately fail-safe.
    return {
      formula,
      tokens: formula.startsWith("=")
        ? [
            {
              span: { start: 0, end: 1 },
              kind: "formula-marker",
              modifiers: [],
            },
          ]
        : [],
      references: [],
      diagnostics: [
        {
          span: { start: 0, end: formula.length },
          severity: "error",
          code: "analysis.internal-error",
          message:
            error instanceof Error
              ? `Formula analysis failed: ${error.message}`
              : "Formula analysis failed.",
        },
      ],
    };
  }
}

function markTokensForQualifier(
  tokens: WorkingToken[],
  qualifier: ReferenceQualifier
): void {
  markTokensByIndexes(tokens, qualifier.workbookTokenIndexes, "workbook-name");
  markTokensByIndexes(tokens, qualifier.sheetTokenIndexes, "sheet-name");
}

function markTokensByIndexes(
  tokens: WorkingToken[],
  indexes: readonly number[],
  kind: FormulaTokenKind
): void {
  for (const index of indexes) {
    const token = tokens[index];
    if (token) token.kind = kind;
  }
}

/**
 * Return the semantic reference at a UTF-16 offset. At a token boundary, the
 * reference immediately to the left wins so a caret after `A2` stays active.
 */
export function findFormulaReferenceAt(
  analysis: FormulaAnalysis,
  offset: number
): FormulaReference | undefined {
  if (!Number.isFinite(offset)) return undefined;
  const containing = analysis.references
    .filter(
      (reference) =>
        reference.span.start <= offset && offset < reference.span.end
    )
    .sort(
      (left, right) =>
        left.span.end - left.span.start - (right.span.end - right.span.start)
    )[0];
  if (containing) return containing;
  return analysis.references.find((reference) => reference.span.end === offset);
}
