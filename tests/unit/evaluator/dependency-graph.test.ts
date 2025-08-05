import { test, expect, describe } from "bun:test";
import { DependencyGraph } from '../../../src/evaluator/dependency-graph';
import type { SimpleCellAddress, SimpleCellRange } from '../../../src/core/types';

describe('DependencyGraph', () => {
  describe('Cell operations', () => {
    test('should add and retrieve cell nodes', () => {
      const graph = new DependencyGraph();
      const address: SimpleCellAddress = { sheet: 0, col: 1, row: 2 };
      
      const key = graph.addCell(address);
      expect(key).toBe('0:1:2');
      
      const nodes = graph.getAllNodes();
      expect(nodes).toHaveLength(1);
      expect(nodes[0]?.type).toBe('cell');
      expect((nodes[0] as any)?.address).toEqual(address);
    });

    test('should handle duplicate cell additions', () => {
      const graph = new DependencyGraph();
      const address: SimpleCellAddress = { sheet: 0, col: 1, row: 2 };
      
      graph.addCell(address);
      graph.addCell(address);
      
      expect(graph.size).toBe(1);
    });

    test('should create correct cell keys', () => {
      expect(DependencyGraph.getCellKey({ sheet: 0, col: 0, row: 0 })).toBe('0:0:0');
      expect(DependencyGraph.getCellKey({ sheet: 1, col: 25, row: 99 })).toBe('1:25:99');
    });
  });

  describe('Range operations', () => {
    test('should add and retrieve range nodes', () => {
      const graph = new DependencyGraph();
      const range: SimpleCellRange = {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 2, row: 2 }
      };
      
      const key = graph.addRange(range);
      expect(key).toBe('0:0:0:2:2');
      
      const nodes = graph.getAllNodes();
      expect(nodes).toHaveLength(1);
      expect(nodes[0]?.type).toBe('range');
      expect((nodes[0] as any)?.range).toEqual(range);
    });

    test('should index cells in range', () => {
      const graph = new DependencyGraph();
      const range: SimpleCellRange = {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 1, row: 1 }
      };
      
      graph.addRange(range);
      
      // Check that all cells in range are indexed
      expect(graph.isCellInRange({ sheet: 0, col: 0, row: 0 })).toBe(true);
      expect(graph.isCellInRange({ sheet: 0, col: 0, row: 1 })).toBe(true);
      expect(graph.isCellInRange({ sheet: 0, col: 1, row: 0 })).toBe(true);
      expect(graph.isCellInRange({ sheet: 0, col: 1, row: 1 })).toBe(true);
      
      // Check cells outside range
      expect(graph.isCellInRange({ sheet: 0, col: 2, row: 0 })).toBe(false);
      expect(graph.isCellInRange({ sheet: 0, col: 0, row: 2 })).toBe(false);
    });

    test('should get ranges containing a cell', () => {
      const graph = new DependencyGraph();
      const range1: SimpleCellRange = {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 2, row: 2 }
      };
      const range2: SimpleCellRange = {
        start: { sheet: 0, col: 1, row: 1 },
        end: { sheet: 0, col: 3, row: 3 }
      };
      
      const key1 = graph.addRange(range1);
      const key2 = graph.addRange(range2);
      
      // Cell in both ranges
      const ranges = graph.getRangesContainingCell({ sheet: 0, col: 1, row: 1 });
      expect(ranges).toHaveLength(2);
      expect(ranges).toContain(key1);
      expect(ranges).toContain(key2);
      
      // Cell in only first range
      const ranges2 = graph.getRangesContainingCell({ sheet: 0, col: 0, row: 0 });
      expect(ranges2).toHaveLength(1);
      expect(ranges2).toContain(key1);
    });

    test('should create correct range keys', () => {
      const range: SimpleCellRange = {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 9, row: 9 }
      };
      expect(DependencyGraph.getRangeKey(range)).toBe('0:0:0:9:9');
    });
  });

  describe('Named expression operations', () => {
    test('should add named expressions', () => {
      const graph = new DependencyGraph();
      
      const globalKey = graph.addNamedExpression('GlobalName');
      expect(globalKey).toBe('name:GlobalName');
      
      const sheetKey = graph.addNamedExpression('SheetName', 1);
      expect(sheetKey).toBe('name:1:SheetName');
      
      expect(graph.size).toBe(2);
    });

    test('should create correct named expression keys', () => {
      expect(DependencyGraph.getNamedExpressionKey('TaxRate')).toBe('name:TaxRate');
      expect(DependencyGraph.getNamedExpressionKey('Discount', 2)).toBe('name:2:Discount');
    });
  });

  describe('Dependency management', () => {
    test('should add dependencies between nodes', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      
      graph.addDependency(a1, b1); // A1 depends on B1
      
      expect(graph.getPrecedents(a1)).toEqual([b1]);
      expect(graph.getDependents(b1)).toEqual([a1]);
    });

    test('should remove dependencies', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      
      graph.addDependency(a1, b1);
      graph.addDependency(a1, c1);
      
      graph.removeDependency(a1, b1);
      
      expect(graph.getPrecedents(a1)).toEqual([c1]);
      expect(graph.getDependents(b1)).toEqual([]);
    });

    test('should clear all dependencies for a node', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      
      graph.addDependency(a1, b1);
      graph.addDependency(a1, c1);
      
      graph.clearDependencies(a1);
      
      expect(graph.getPrecedents(a1)).toEqual([]);
      expect(graph.getDependents(b1)).toEqual([]);
      expect(graph.getDependents(c1)).toEqual([]);
    });

    test('should throw error when adding dependency between non-existent nodes', () => {
      const graph = new DependencyGraph();
      
      expect(() => {
        graph.addDependency('0:0:0', '0:1:0');
      }).toThrow('Cannot add dependency between non-existent nodes');
    });
  });

  describe('Transitive dependencies', () => {
    test('should get transitive dependents', () => {
      const graph = new DependencyGraph();
      
      // Create chain: D1 <- C1 <- B1 <- A1
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      const d1 = graph.addCell({ sheet: 0, col: 3, row: 0 });
      
      graph.addDependency(b1, a1); // B1 depends on A1
      graph.addDependency(c1, b1); // C1 depends on B1
      graph.addDependency(d1, c1); // D1 depends on C1
      
      const dependents = graph.getTransitiveDependents(a1);
      expect(dependents.size).toBe(3);
      expect(dependents.has(b1)).toBe(true);
      expect(dependents.has(c1)).toBe(true);
      expect(dependents.has(d1)).toBe(true);
    });

    test('should get transitive precedents', () => {
      const graph = new DependencyGraph();
      
      // Create chain: D1 -> C1 -> B1 -> A1
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      const d1 = graph.addCell({ sheet: 0, col: 3, row: 0 });
      
      graph.addDependency(d1, c1); // D1 depends on C1
      graph.addDependency(c1, b1); // C1 depends on B1
      graph.addDependency(b1, a1); // B1 depends on A1
      
      const precedents = graph.getTransitivePrecedents(d1);
      expect(precedents.size).toBe(3);
      expect(precedents.has(c1)).toBe(true);
      expect(precedents.has(b1)).toBe(true);
      expect(precedents.has(a1)).toBe(true);
    });

    test('should handle diamond dependencies', () => {
      const graph = new DependencyGraph();
      
      //    B1
      //   /  \
      // A1    D1
      //   \  /
      //    C1
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      const d1 = graph.addCell({ sheet: 0, col: 3, row: 0 });
      
      graph.addDependency(b1, a1);
      graph.addDependency(c1, a1);
      graph.addDependency(d1, b1);
      graph.addDependency(d1, c1);
      
      const dependents = graph.getTransitiveDependents(a1);
      expect(dependents.size).toBe(3);
      expect(dependents.has(b1)).toBe(true);
      expect(dependents.has(c1)).toBe(true);
      expect(dependents.has(d1)).toBe(true);
    });
  });

  describe('Cycle detection', () => {
    test('should detect simple cycle', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      
      graph.addDependency(a1, b1); // A1 depends on B1
      graph.addDependency(b1, a1); // B1 depends on A1 - cycle!
      
      const result = graph.detectCycles();
      expect(result.hasCycle).toBe(true);
      expect(result.cycle).toBeDefined();
      expect(result.cycle).toContain(a1);
      expect(result.cycle).toContain(b1);
    });

    test('should detect longer cycle', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      const d1 = graph.addCell({ sheet: 0, col: 3, row: 0 });
      
      graph.addDependency(a1, b1); // A1 -> B1
      graph.addDependency(b1, c1); // B1 -> C1
      graph.addDependency(c1, d1); // C1 -> D1
      graph.addDependency(d1, a1); // D1 -> A1 - cycle!
      
      const result = graph.detectCycles();
      expect(result.hasCycle).toBe(true);
      expect(result.cycle).toBeDefined();
      expect(result.cycle?.length).toBeGreaterThanOrEqual(4);
    });

    test('should not detect cycle in acyclic graph', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      
      graph.addDependency(a1, b1);
      graph.addDependency(a1, c1);
      graph.addDependency(b1, c1);
      
      const result = graph.detectCycles();
      expect(result.hasCycle).toBe(false);
      expect(result.cycle).toBeUndefined();
    });
  });

  describe('Topological sort', () => {
    test('should sort nodes in dependency order', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      
      // No dependencies between nodes - all independent
      
      const sorted = graph.topologicalSort();
      expect(sorted).toBeDefined();
      expect(sorted).toHaveLength(3);
      
      // All nodes should be in the result
      expect(sorted).toContain(a1);
      expect(sorted).toContain(b1);
      expect(sorted).toContain(c1);
    });

    test('should return null for cyclic graph', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      
      graph.addDependency(a1, b1);
      graph.addDependency(b1, a1);
      
      const sorted = graph.topologicalSort();
      expect(sorted).toBeNull();
    });

    test('should handle disconnected components', () => {
      const graph = new DependencyGraph();
      
      // Component 1: independent nodes
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      
      // Component 2: independent nodes
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      const d1 = graph.addCell({ sheet: 0, col: 3, row: 0 });
      
      const sorted = graph.topologicalSort();
      expect(sorted).toBeDefined();
      expect(sorted).toHaveLength(4);
      
      // All nodes should be present
      expect(sorted).toContain(a1);
      expect(sorted).toContain(b1);
      expect(sorted).toContain(c1);
      expect(sorted).toContain(d1);
    });
  });

  describe('Node removal', () => {
    test('should remove node and its dependencies', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      const c1 = graph.addCell({ sheet: 0, col: 2, row: 0 });
      
      graph.addDependency(a1, b1);
      graph.addDependency(c1, b1);
      
      graph.removeNode(b1);
      
      expect(graph.size).toBe(2);
      expect(graph.getPrecedents(a1)).toEqual([]);
      expect(graph.getPrecedents(c1)).toEqual([]);
    });

    test('should remove range and update index', () => {
      const graph = new DependencyGraph();
      
      const range: SimpleCellRange = {
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 1, row: 1 }
      };
      
      const key = graph.addRange(range);
      expect(graph.isCellInRange({ sheet: 0, col: 0, row: 0 })).toBe(true);
      
      graph.removeNode(key);
      
      expect(graph.size).toBe(0);
      expect(graph.isCellInRange({ sheet: 0, col: 0, row: 0 })).toBe(false);
    });
  });

  describe('Graph management', () => {
    test('should clear entire graph', () => {
      const graph = new DependencyGraph();
      
      graph.addCell({ sheet: 0, col: 0, row: 0 });
      graph.addCell({ sheet: 0, col: 1, row: 0 });
      graph.addRange({
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 5, row: 5 }
      });
      
      expect(graph.size).toBe(3);
      
      graph.clear();
      
      expect(graph.size).toBe(0);
      expect(graph.getAllNodes()).toEqual([]);
    });

    test('should provide string representation', () => {
      const graph = new DependencyGraph();
      
      const a1 = graph.addCell({ sheet: 0, col: 0, row: 0 });
      const b1 = graph.addCell({ sheet: 0, col: 1, row: 0 });
      
      graph.addDependency(a1, b1);
      
      const str = graph.toString();
      expect(str).toContain(a1);
      expect(str).toContain(b1);
      expect(str).toContain('depends on');
      expect(str).toContain('used by');
    });
  });
});