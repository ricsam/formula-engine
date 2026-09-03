import { FormulaError, type SerializedCellValue } from "../core/types";
import type {
  ASTNode,
  BinaryOpNode,
  FunctionNode,
  UnaryOpNode,
} from "../parser/ast";
import { astToString } from "../parser/formatter";
import {
  getOperatorAssociativity,
  getOperatorPrecedence,
} from "../parser/grammar";
import { Lexer } from "../parser/lexer";
import { ParseError, parseFormula } from "../parser/parser";
import type { SourceSpan } from "./formula-analysis";

/** The layout produced by the formula formatter. */
export type FormulaFormatStyle = "compact" | "pretty";

export interface FormulaFormatOptions {
  /** Defaults to `compact`. */
  style?: FormulaFormatStyle;
  /** One indentation level in pretty output. Defaults to two spaces. */
  indent?: string;
}

export interface FormulaFormatError {
  name: string;
  message: string;
  /** A half-open UTF-16 range in the original formula, when available. */
  position?: SourceSpan;
}

export type FormulaFormatResult =
  | {
      ok: true;
      formula: string;
      changed: boolean;
    }
  | {
      ok: false;
      /** Invalid input is returned unchanged. */
      formula: string;
      changed: false;
      error: FormulaFormatError;
    };

interface RenderContext {
  style: FormulaFormatStyle;
  indent: string;
}

const KNOWN_FORMULA_ERRORS = new Set<string>(Object.values(FormulaError));

function validateOptions(options: FormulaFormatOptions): RenderContext {
  const style = options.style ?? "compact";
  const indent = options.indent ?? "  ";

  if (style !== "compact" && style !== "pretty") {
    throw new TypeError(`Unknown formula format style: ${String(style)}`);
  }
  if (indent.includes("\n") || indent.includes("\r")) {
    throw new TypeError("Formula indentation cannot contain a newline.");
  }

  return { style, indent };
}

function needsParentheses(
  child: ASTNode,
  parent: BinaryOpNode,
  position: "left" | "right"
): boolean {
  if (child.type !== "binary-op") return false;

  const childPrecedence = getOperatorPrecedence(child.operator);
  const parentPrecedence = getOperatorPrecedence(parent.operator);
  if (childPrecedence < parentPrecedence) return true;
  if (childPrecedence > parentPrecedence) return false;

  const associativity = getOperatorAssociativity(parent.operator);
  return (
    (associativity === "left" && position === "right") ||
    (associativity === "right" && position === "left")
  );
}

function renderFunction(
  node: FunctionNode,
  context: RenderContext,
  depth: number
): string {
  if (node.args.length === 0) return `${node.name}()`;

  if (context.style === "compact") {
    return `${node.name}(${node.args
      .map((argument) => renderNode(argument, context, depth))
      .join(", ")})`;
  }

  const childIndent = context.indent.repeat(depth + 1);
  const closingIndent = context.indent.repeat(depth);
  const argumentsText = node.args
    .map(
      (argument) => `${childIndent}${renderNode(argument, context, depth + 1)}`
    )
    .join(",\n");
  return `${node.name}(\n${argumentsText}\n${closingIndent})`;
}

function renderUnary(
  node: UnaryOpNode,
  context: RenderContext,
  depth: number
): string {
  const renderedOperand = renderNode(node.operand, context, depth);
  const operand =
    node.operand.type === "binary-op"
      ? `(${renderedOperand})`
      : renderedOperand;
  return node.operator === "%" ? `${operand}%` : `${node.operator}${operand}`;
}

function renderBinary(
  node: BinaryOpNode,
  context: RenderContext,
  depth: number
): string {
  const renderedLeft = renderNode(node.left, context, depth);
  const renderedRight = renderNode(node.right, context, depth);
  const left = needsParentheses(node.left, node, "left")
    ? `(${renderedLeft})`
    : renderedLeft;
  const right = needsParentheses(node.right, node, "right")
    ? `(${renderedRight})`
    : renderedRight;
  const operator =
    context.style === "pretty" ? ` ${node.operator} ` : node.operator;
  return `${left}${operator}${right}`;
}

function renderArray(
  node: Extract<ASTNode, { type: "array" }>,
  context: RenderContext,
  depth: number
): string {
  if (context.style === "compact") {
    const rows = node.elements.map((row) =>
      row.map((cell) => renderNode(cell, context, depth)).join(", ")
    );
    return `{${rows.join("; ")}}`;
  }

  if (node.elements.length === 0) return "{}";
  const rowIndent = context.indent.repeat(depth + 1);
  const closingIndent = context.indent.repeat(depth);
  const rows = node.elements.map(
    (row) =>
      `${rowIndent}${row
        .map((cell) => renderNode(cell, context, depth + 1))
        .join(", ")}`
  );
  return `{\n${rows.join(";\n")}\n${closingIndent}}`;
}

function renderNode(
  node: ASTNode,
  context: RenderContext,
  depth: number
): string {
  switch (node.type) {
    case "function":
      return renderFunction(node, context, depth);
    case "unary-op":
      return renderUnary(node, context, depth);
    case "binary-op":
      return renderBinary(node, context, depth);
    case "array":
      return renderArray(node, context, depth);
    default:
      return astToString(node);
  }
}

function hasClosingStringQuote(source: string): boolean {
  let index = 1;
  while (index < source.length) {
    if (source[index] !== '"') {
      index++;
      continue;
    }
    if (source[index + 1] === '"') {
      index += 2;
      continue;
    }
    return index === source.length - 1;
  }
  return false;
}

function assertLexicallyValidFormula(
  formula: string,
  body: string,
  bodyOffset: number
): void {
  if (formula.startsWith("=") && formula.slice(1).trim().length === 0) {
    throw new ParseError("Expected an expression after '='.", {
      start: formula.length,
      end: formula.length,
    });
  }

  for (const token of new Lexer(body).tokenize()) {
    const position = {
      start: token.position.start + bodyOffset,
      end: token.position.end + bodyOffset,
    };
    if (token.type === "ERROR" && !KNOWN_FORMULA_ERRORS.has(token.value)) {
      throw new ParseError(
        `Invalid token ${JSON.stringify(
          formula.slice(position.start, position.end)
        )}.`,
        position
      );
    }
    if (
      token.type === "STRING" &&
      !hasClosingStringQuote(
        body.slice(token.position.start, token.position.end)
      )
    ) {
      throw new ParseError("Unterminated string literal.", position);
    }
  }
}

function parseInputBody(body: string, bodyOffset: number): ASTNode {
  try {
    return parseFormula(body);
  } catch (error) {
    if (error instanceof ParseError && error.position && bodyOffset !== 0) {
      throw new ParseError(error.message, {
        start: error.position.start + bodyOffset,
        end: error.position.end + bodyOffset,
      });
    }
    throw error;
  }
}

/**
 * Format a formula body or a formula beginning with `=`.
 *
 * This strict form throws when the input is not syntactically valid. Editor
 * integrations should normally use {@link tryFormatFormula} so a partially
 * typed formula can be retained verbatim.
 */
export function formatFormula(
  formula: string,
  options: FormulaFormatOptions = {}
): string {
  const context = validateOptions(options);
  const hasFormulaMarker = formula.startsWith("=");
  const body = hasFormulaMarker ? formula.slice(1) : formula;
  const bodyOffset = hasFormulaMarker ? 1 : 0;
  assertLexicallyValidFormula(formula, body, bodyOffset);
  const ast = parseInputBody(body, bodyOffset);
  const formattedBody = renderNode(ast, context, 0);

  // Refuse to return output if a formatter regression ever changes the AST.
  const reparsed = parseFormula(formattedBody);
  if (astToString(reparsed) !== astToString(ast)) {
    throw new Error("Formula formatting changed the parsed expression.");
  }

  return hasFormulaMarker ? `=${formattedBody}` : formattedBody;
}

/**
 * Format a formula without throwing. Invalid input is returned unchanged.
 */
export function tryFormatFormula(
  formula: string,
  options: FormulaFormatOptions = {}
): FormulaFormatResult {
  try {
    const formatted = formatFormula(formula, options);
    return {
      ok: true,
      formula: formatted,
      changed: formatted !== formula,
    };
  } catch (error) {
    const knownError = error instanceof Error ? error : undefined;
    const position =
      error instanceof ParseError && error.position
        ? { ...error.position }
        : undefined;
    return {
      ok: false,
      formula,
      changed: false,
      error: {
        name: knownError?.name ?? "Error",
        message: knownError?.message ?? String(error),
        ...(position ? { position } : {}),
      },
    };
  }
}

/** @internal Normalize formula cell content without touching text or invalid input. */
export function normalizeFormulaCellContent(
  content: SerializedCellValue
): SerializedCellValue {
  if (typeof content !== "string" || !content.startsWith("=")) {
    return content;
  }

  const result = tryFormatFormula(content, { style: "compact" });
  return result.ok ? result.formula : content;
}
