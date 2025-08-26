import type { ASTNode } from "../parser/ast";

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
