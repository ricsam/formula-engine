import type {
  ASTNode,
  FunctionNode,
  NamedExpressionNode,
} from "../parser/ast";

/**
 * Visitor function type for AST traversal
 */
export type ASTVisitor<T = void> = (node: ASTNode, parent?: ASTNode) => T;

/**
 * Traverses an AST node and all its children, calling the visitor function for each node
 * @param node - The AST node to traverse
 * @param visitor - Function called for each node during traversal
 * @param parent - The parent node (used internally for recursion)
 */
export function traverseAST<T = void>(
  node: ASTNode,
  visitor: ASTVisitor<T>,
  parent?: ASTNode
): void {
  // Visit the current node
  visitor(node, parent);

  // Recursively traverse children based on node type
  switch (node.type) {
    case "binary-op":
      traverseAST(node.left, visitor, node);
      traverseAST(node.right, visitor, node);
      break;

    case "unary-op":
      traverseAST(node.operand, visitor, node);
      break;

    case "function":
      node.args.forEach((arg) => traverseAST(arg, visitor, node));
      break;

    case "array":
      node.elements.forEach((row) =>
        row.forEach((element) => traverseAST(element, visitor, node))
      );
      break;

    case "range":
      // Range nodes don't have start/end child nodes in this AST structure
      // They have a range property with coordinates
      break;

    case "3d-range":
      traverseAST(node.reference, visitor, node);
      break;

    case "structured-reference":
    case "reference":
    case "value":
    case "named-expression":
    case "error":
    case "empty":
    case "infinity":
      // These are leaf nodes, no children to traverse
      break;

    default:
      // Handle any unknown node types gracefully
      console.warn(`Unknown AST node type: ${(node as any).type}`);
      break;
  }
}

/**
 * Finds all nodes of a specific type in an AST
 * @param node - The root AST node to search
 * @param nodeType - The type of nodes to find
 * @returns Array of nodes matching the specified type
 */
export function findNodesByType<T extends ASTNode["type"]>(
  node: ASTNode,
  nodeType: T
): Extract<ASTNode, { type: T }>[] {
  const results: Extract<ASTNode, { type: T }>[] = [];

  traverseAST(node, (currentNode) => {
    if (currentNode.type === nodeType) {
      results.push(currentNode as Extract<ASTNode, { type: T }>);
    }
  });

  return results;
}

/**
 * Transforms an AST by applying a transformation function to each node
 * @param node - The AST node to transform
 * @param transformer - Function that transforms each node
 * @returns The transformed AST
 */
export function transformAST(
  node: ASTNode,
  transformer: (node: ASTNode, parent?: ASTNode) => ASTNode
): ASTNode {
  const transform = (currentNode: ASTNode, parent?: ASTNode): ASTNode => {
    // First transform children
    let transformedNode: ASTNode;

    switch (currentNode.type) {
      case "binary-op":
        transformedNode = {
          ...currentNode,
          left: transform(currentNode.left, currentNode),
          right: transform(currentNode.right, currentNode),
        };
        break;

      case "unary-op":
        transformedNode = {
          ...currentNode,
          operand: transform(currentNode.operand, currentNode),
        };
        break;

      case "function":
        transformedNode = {
          ...currentNode,
          args: currentNode.args.map((arg) => transform(arg, currentNode)),
        };
        break;

      case "array":
        transformedNode = {
          ...currentNode,
          elements: currentNode.elements.map((row) =>
            row.map((element) => transform(element, currentNode))
          ),
        };
        break;

      case "3d-range":
        transformedNode = {
          ...currentNode,
          reference: transform(currentNode.reference, currentNode) as typeof currentNode.reference,
        };
        break;

      default:
        // Leaf nodes or unknown types - return as is
        transformedNode = { ...currentNode };
        break;
    }

    // Then apply the transformer to the current node
    return transformer(transformedNode, parent);
  };

  return transform(node);
}

/**
 * LET is the only function that binds names. Its name arguments and the
 * references to them parse as `named-expression` nodes, exactly like references
 * to defined names, so tooling that inspects names must know LET's scoping
 * rules to tell the two apart.
 */
export const LET_FUNCTION_NAME = "LET";

/** LET names match without regard to case, like Excel names. */
export function normalizeLetName(name: string): string {
  return name.toUpperCase();
}

export function isLetCall(
  node: ASTNode
): node is FunctionNode & { name: typeof LET_FUNCTION_NAME } {
  return node.type === "function" && node.name === LET_FUNCTION_NAME;
}

/**
 * The name declared by a LET name argument: a plain identifier that parses as
 * an unqualified named expression. Anything else (a cell reference such as
 * `x1`, a sheet- or workbook-qualified name, a literal, an expression) is not a
 * valid LET name.
 */
export function getLetDeclaredName(node: ASTNode): string | undefined {
  return node.type === "named-expression" &&
    node.sheetName === undefined &&
    node.workbookName === undefined
    ? node.name
    : undefined;
}

/** One name declared by a LET call. */
export interface LetBinding {
  /** The LET call that declares the name. */
  letCall: FunctionNode;
  /** Index of the name argument in `letCall.args`. */
  argIndex: number;
  /** The name argument itself. */
  declaration: NamedExpressionNode;
  /** The binding of the same name that this one shadows, if any. */
  shadows?: LetBinding;
}

/** The LET bindings in scope, keyed by {@link normalizeLetName}. */
export type LetScope = ReadonlyMap<string, LetBinding>;

/**
 * What a `named-expression` node means at its position:
 * - `declaration`: the name argument of a LET pair;
 * - `local`: a reference to a LET name in scope;
 * - `defined`: a reference to a defined name (named expression). Sheet- or
 *   workbook-qualified names are always defined names.
 */
export type NameRole =
  | { kind: "declaration"; binding: LetBinding }
  | { kind: "local"; binding: LetBinding }
  | { kind: "defined" };

/**
 * Rebuild an AST, letting `mapName` replace each `named-expression` node with
 * knowledge of LET scoping. Other nodes are copied with their mapped children.
 *
 * Scoping follows LET's evaluation rules: a name is visible in the values of
 * later pairs and in the final calculation of its LET, not in its own value;
 * nested LETs shadow outer names. The walk is lenient so that tooling can run
 * on LET calls that would fail to evaluate: an even-indexed argument followed
 * by another argument is a name position, a name position holding something
 * other than a plain name declares nothing and is walked as an expression, and
 * a repeated name shadows the earlier one.
 */
export function mapNamesWithLetScopes(
  node: ASTNode,
  mapName: (
    node: NamedExpressionNode,
    role: NameRole,
    scope: LetScope
  ) => ASTNode,
  scope: LetScope = new Map()
): ASTNode {
  const map = (child: ASTNode, childScope: LetScope = scope) =>
    mapNamesWithLetScopes(child, mapName, childScope);

  switch (node.type) {
    case "named-expression": {
      const binding =
        getLetDeclaredName(node) === undefined
          ? undefined
          : scope.get(normalizeLetName(node.name));
      return mapName(
        node,
        binding ? { kind: "local", binding } : { kind: "defined" },
        scope
      );
    }

    case "function": {
      if (!isLetCall(node)) {
        return { ...node, args: node.args.map((arg) => map(arg)) };
      }

      const args: ASTNode[] = [];
      let letScope = scope;
      for (let index = 0; index < node.args.length; index += 2) {
        const nameArgument = node.args[index]!;
        if (index === node.args.length - 1) {
          // The final calculation sees every name of this LET.
          args.push(map(nameArgument, letScope));
          break;
        }

        const valueArgument = node.args[index + 1]!;
        const name = getLetDeclaredName(nameArgument);
        if (name === undefined) {
          args.push(map(nameArgument, letScope), map(valueArgument, letScope));
          continue;
        }

        const key = normalizeLetName(name);
        const binding: LetBinding = {
          letCall: node,
          argIndex: index,
          declaration: nameArgument as NamedExpressionNode,
          shadows: letScope.get(key),
        };
        args.push(
          mapName(
            nameArgument as NamedExpressionNode,
            { kind: "declaration", binding },
            letScope
          ),
          // A value sees the names declared before its own.
          map(valueArgument, letScope)
        );
        letScope = new Map(letScope).set(key, binding);
      }
      return { ...node, args };
    }

    case "binary-op":
      return { ...node, left: map(node.left), right: map(node.right) };

    case "unary-op":
      return { ...node, operand: map(node.operand) };

    case "array":
      return {
        ...node,
        elements: node.elements.map((row) => row.map((element) => map(element))),
      };

    case "3d-range":
      return {
        ...node,
        reference: map(node.reference) as typeof node.reference,
      };

    default:
      return node;
  }
}

/**
 * Visit each `named-expression` node with its LET-scoped role. See
 * {@link mapNamesWithLetScopes} for the scoping rules.
 */
export function forEachNameWithLetScopes(
  node: ASTNode,
  visit: (node: NamedExpressionNode, role: NameRole, scope: LetScope) => void
): void {
  mapNamesWithLetScopes(node, (name, role, scope) => {
    visit(name, role, scope);
    return name;
  });
}
