import { test, expect, describe } from "bun:test";
import { parseFormula } from "../../../src/parser/parser";
import { traverseAST, findNodesByType, transformAST } from "../../../src/core/ast-traverser";
import type { ASTNode } from "../../../src/parser/ast";

describe("AST Traverser", () => {
  describe("traverseAST", () => {
    test("should visit all nodes in a simple formula", () => {
      const ast = parseFormula("A1+B1");
      const visitedNodes: string[] = [];

      traverseAST(ast, (node) => {
        visitedNodes.push(node.type);
      });

      expect(visitedNodes).toEqual([
        "binary-op",
        "reference",
        "reference"
      ]);
    });

    test("should visit all nodes in a complex formula", () => {
      const ast = parseFormula("SUM(A1:B2, C1*2)");
      const visitedNodes: string[] = [];

      traverseAST(ast, (node) => {
        visitedNodes.push(node.type);
      });

      expect(visitedNodes).toEqual([
        "function",
        "range",
        "binary-op",
        "reference",
        "value"
      ]);
    });

    test("should visit structured reference nodes", () => {
      const ast = parseFormula("SUM(Table1[Price])");
      const visitedNodes: string[] = [];

      traverseAST(ast, (node) => {
        visitedNodes.push(node.type);
      });

      expect(visitedNodes).toEqual([
        "function",
        "structured-reference"
      ]);
    });

    test("should provide parent node in visitor", () => {
      const ast = parseFormula("A1+B1");
      const nodeParents: Array<{ nodeType: string; parentType?: string }> = [];

      traverseAST(ast, (node, parent) => {
        nodeParents.push({
          nodeType: node.type,
          parentType: parent?.type
        });
      });

      expect(nodeParents).toEqual([
        { nodeType: "binary-op", parentType: undefined },
        { nodeType: "reference", parentType: "binary-op" },
        { nodeType: "reference", parentType: "binary-op" }
      ]);
    });
  });

  describe("findNodesByType", () => {
    test("should find all cell references", () => {
      const ast = parseFormula("A1+B1*C1");
      const cellRefs = findNodesByType(ast, "reference");

      expect(cellRefs).toHaveLength(3);
      // Reference nodes have address property, not cell property
      expect(cellRefs.map(ref => `${String.fromCharCode(65 + ref.address.colIndex)}${ref.address.rowIndex + 1}`)).toEqual(["A1", "B1", "C1"]);
    });

    test("should find structured references", () => {
      const ast = parseFormula("SUM(Table1[Price], Table2[Quantity])");
      const structuredRefs = findNodesByType(ast, "structured-reference");

      expect(structuredRefs).toHaveLength(2);
      expect(structuredRefs.map(ref => ref.tableName)).toEqual(["Table1", "Table2"]);
      
      // Check the actual structure - structured references have cols property
      expect(structuredRefs.map(ref => ref.cols?.startCol)).toEqual(["Price", "Quantity"]);
    });

    test("should find function calls", () => {
      const ast = parseFormula("SUM(A1:B2) + INDEX(C1:D4, 1, 1)");
      const functionCalls = findNodesByType(ast, "function");

      expect(functionCalls).toHaveLength(2);
      expect(functionCalls.map(fn => fn.name)).toEqual(["SUM", "INDEX"]);
    });

    test("should return empty array when no nodes match", () => {
      const ast = parseFormula("42");
      const cellRefs = findNodesByType(ast, "reference");

      expect(cellRefs).toHaveLength(0);
    });
  });

  describe("transformAST", () => {
    test("should transform cell references", () => {
      const ast = parseFormula("A1+B1");
      
      const transformed = transformAST(ast, (node) => {
        if (node.type === "reference") {
          return {
            ...node,
            sheetName: "Sheet1"
          };
        }
        return node;
      });

      const cellRefs = findNodesByType(transformed, "reference");
      expect(cellRefs.map(ref => ref.sheetName)).toEqual(["Sheet1", "Sheet1"]);
    });

    test("should transform structured references", () => {
      const ast = parseFormula("SUM(OldTable[Price])");
      
      const transformed = transformAST(ast, (node) => {
        if (node.type === "structured-reference" && node.tableName === "OldTable") {
          return {
            ...node,
            tableName: "NewTable"
          };
        }
        return node;
      });

      const structuredRefs = findNodesByType(transformed, "structured-reference");
      expect(structuredRefs[0]?.tableName).toBe("NewTable");
      expect(structuredRefs[0]?.cols?.startCol).toBe("Price");
    });

    test("should preserve structure while transforming", () => {
      const ast = parseFormula("SUM(A1:B2, C1*2)");
      
      const transformed = transformAST(ast, (node) => {
        if (node.type === "value" && node.value && typeof node.value === "object" && node.value.type === "number") {
          return {
            ...node,
            value: {
              ...node.value,
              value: node.value.value * 10
            }
          };
        }
        return node;
      });

      const numbers = findNodesByType(transformed, "value");
      const numberValues = numbers.filter(n => n.value && typeof n.value === "object" && n.value.type === "number");
      
      expect(numberValues.length).toBeGreaterThan(0);
      const firstNumber = numberValues[0];
      expect(firstNumber?.value).toBeDefined();
      if (firstNumber?.value && typeof firstNumber.value === "object" && "value" in firstNumber.value) {
        expect(firstNumber.value.value).toBe(20); // 2 * 10
      }
    });

    test("should handle nested transformations", () => {
      const ast = parseFormula("SUM(Table1[Price] * 1.1)");
      
      const transformed = transformAST(ast, (node) => {
        if (node.type === "structured-reference") {
          return {
            ...node,
            tableName: "Products"
          };
        }
        if (node.type === "value" && node.value && typeof node.value === "object" && node.value.type === "number" && node.value.value === 1.1) {
          return {
            ...node,
            value: {
              ...node.value,
              value: 1.2
            }
          };
        }
        return node;
      });

      const structuredRefs = findNodesByType(transformed, "structured-reference");
      const numbers = findNodesByType(transformed, "value");
      const numberValues = numbers.filter(n => n.value && typeof n.value === "object" && n.value.type === "number");
      
      expect(structuredRefs[0]?.tableName).toBe("Products");
      const firstNumber = numberValues[0];
      if (firstNumber?.value && typeof firstNumber.value === "object" && "value" in firstNumber.value) {
        expect(firstNumber.value.value).toBe(1.2);
      }
    });
  });
});
