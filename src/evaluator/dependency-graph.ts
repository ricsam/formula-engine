/**
 * Dependency tracking system for FormulaEngine
 * Manages the directed acyclic graph (DAG) of cell dependencies
 */

import type { SimpleCellAddress, SimpleCellRange } from '../core/types';

/**
 * Node in the dependency graph
 */
export interface DependencyNode {
  id: string;
  type: 'cell' | 'range' | 'named-expression';
  address?: SimpleCellAddress;
  range?: SimpleCellRange;
  name?: string;
  scope?: number;
}

/**
 * Edge in the dependency graph
 */
export interface DependencyEdge {
  from: string;
  to: string;
}

/**
 * Dependency information for a node
 */
export interface DependencyInfo {
  precedents: Set<string>; // Nodes this depends on
  dependents: Set<string>; // Nodes that depend on this
}

/**
 * Result of cycle detection
 */
export interface CycleDetectionResult {
  hasCycle: boolean;
  cycle?: string[];
}

/**
 * Dependency graph implementation
 */
export class DependencyGraph {
  private nodes: Map<string, DependencyNode> = new Map();
  private dependencies: Map<string, DependencyInfo> = new Map();
  private rangeIndex: Map<string, Set<string>> = new Map(); // Maps cells to ranges containing them
  private largeRanges: Array<{key: string, range: SimpleCellRange, size: number}> = []; // Sparse index for large ranges
  
  /**
   * Creates a unique key for a cell address
   */
  static getCellKey(address: SimpleCellAddress): string {
    return `${address.sheet}:${address.col}:${address.row}`;
  }
  
  /**
   * Creates a unique key for a range
   */
  static getRangeKey(range: SimpleCellRange): string {
    return `${range.start.sheet}:${range.start.col}:${range.start.row}:${range.end.col}:${range.end.row}`;
  }
  
  /**
   * Creates a unique key for a named expression
   */
  static getNamedExpressionKey(name: string, scope?: number): string {
    return scope === undefined ? `name:${name}` : `name:${scope}:${name}`;
  }
  
  /**
   * Adds a cell node to the graph
   */
  addCell(address: SimpleCellAddress): string {
    const key = DependencyGraph.getCellKey(address);
    
    if (!this.nodes.has(key)) {
      this.nodes.set(key, {
        id: key,
        type: 'cell',
        address
      });
      
      this.dependencies.set(key, {
        precedents: new Set(),
        dependents: new Set()
      });
    }
    
    return key;
  }
  
  /**
   * Adds a range node to the graph
   */
  addRange(range: SimpleCellRange): string {
    const key = DependencyGraph.getRangeKey(range);
    
    if (!this.nodes.has(key)) {
      this.nodes.set(key, {
        id: key,
        type: 'range',
        range
      });
      
      this.dependencies.set(key, {
        precedents: new Set(),
        dependents: new Set()
      });
      
      // Index all cells in the range using optimized approach
      // For large ranges, we use sparse indexing to avoid O(n²) complexity
      const rangeSize = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
      
      if (rangeSize > 1000) {
        // For large ranges, use sparse indexing
        // We'll track the range boundaries and check containment on demand
        this.largeRanges.push({
          key,
          range,
          size: rangeSize
        });
      } else {
        // For smaller ranges, use direct indexing
        for (let row = range.start.row; row <= range.end.row; row++) {
          for (let col = range.start.col; col <= range.end.col; col++) {
            const cellKey = DependencyGraph.getCellKey({
              sheet: range.start.sheet,
              col,
              row
            });
            
            if (!this.rangeIndex.has(cellKey)) {
              this.rangeIndex.set(cellKey, new Set());
            }
            this.rangeIndex.get(cellKey)!.add(key);
          }
        }
      }
    }
    
    return key;
  }
  
  /**
   * Adds a named expression node to the graph
   */
  addNamedExpression(name: string, scope?: number): string {
    const key = DependencyGraph.getNamedExpressionKey(name, scope);
    
    if (!this.nodes.has(key)) {
      this.nodes.set(key, {
        id: key,
        type: 'named-expression',
        name,
        scope
      });
      
      this.dependencies.set(key, {
        precedents: new Set(),
        dependents: new Set()
      });
    }
    
    return key;
  }
  
  /**
   * Removes a node from the graph
   */
  removeNode(key: string): void {
    const deps = this.dependencies.get(key);
    if (!deps) return;
    
    // Remove this node from its precedents' dependents
    for (const precedent of deps.precedents) {
      const precedentDeps = this.dependencies.get(precedent);
      if (precedentDeps) {
        precedentDeps.dependents.delete(key);
      }
    }
    
    // Remove this node from its dependents' precedents
    for (const dependent of deps.dependents) {
      const dependentDeps = this.dependencies.get(dependent);
      if (dependentDeps) {
        dependentDeps.precedents.delete(key);
      }
    }
    
    // Remove from range index if it's a range
    const node = this.nodes.get(key);
    if (node?.type === 'range' && node.range) {
      const range = node.range;
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.col; col <= range.end.col; col++) {
          const cellKey = DependencyGraph.getCellKey({
            sheet: range.start.sheet,
            col,
            row
          });
          
          const ranges = this.rangeIndex.get(cellKey);
          if (ranges) {
            ranges.delete(key);
            if (ranges.size === 0) {
              this.rangeIndex.delete(cellKey);
            }
          }
        }
      }
    }
    
    this.nodes.delete(key);
    this.dependencies.delete(key);
  }
  
  /**
   * Adds a dependency edge
   */
  addDependency(fromKey: string, toKey: string): void {
    const fromDeps = this.dependencies.get(fromKey);
    const toDeps = this.dependencies.get(toKey);
    
    if (!fromDeps || !toDeps) {
      throw new Error('Cannot add dependency between non-existent nodes');
    }
    
    fromDeps.precedents.add(toKey);
    toDeps.dependents.add(fromKey);
  }
  
  /**
   * Removes a dependency edge
   */
  removeDependency(fromKey: string, toKey: string): void {
    const fromDeps = this.dependencies.get(fromKey);
    const toDeps = this.dependencies.get(toKey);
    
    if (fromDeps) {
      fromDeps.precedents.delete(toKey);
    }
    
    if (toDeps) {
      toDeps.dependents.delete(fromKey);
    }
  }
  
  /**
   * Clears all dependencies for a node
   */
  clearDependencies(key: string): void {
    const deps = this.dependencies.get(key);
    if (!deps) return;
    
    // Remove from precedents
    for (const precedent of deps.precedents) {
      const precedentDeps = this.dependencies.get(precedent);
      if (precedentDeps) {
        precedentDeps.dependents.delete(key);
      }
    }
    
    deps.precedents.clear();
  }
  
  /**
   * Gets direct precedents of a node
   */
  getPrecedents(key: string): string[] {
    const deps = this.dependencies.get(key);
    return deps ? Array.from(deps.precedents) : [];
  }
  
  /**
   * Gets direct dependents of a node
   */
  getDependents(key: string): string[] {
    const deps = this.dependencies.get(key);
    return deps ? Array.from(deps.dependents) : [];
  }
  
  /**
   * Gets all transitive dependents (cells affected by changes to this node)
   */
  getTransitiveDependents(key: string): Set<string> {
    const visited = new Set<string>();
    const queue = [key];
    
    while (queue.length > 0) {
      const current = queue.shift()!;
      
      if (visited.has(current)) continue;
      visited.add(current);
      
      const deps = this.dependencies.get(current);
      if (deps) {
        for (const dependent of deps.dependents) {
          queue.push(dependent);
        }
      }
    }
    
    visited.delete(key); // Don't include the starting node
    return visited;
  }
  
  /**
   * Gets all transitive precedents (cells this node depends on)
   */
  getTransitivePrecedents(key: string): Set<string> {
    const visited = new Set<string>();
    const queue = [key];
    
    while (queue.length > 0) {
      const current = queue.shift()!;
      
      if (visited.has(current)) continue;
      visited.add(current);
      
      const deps = this.dependencies.get(current);
      if (deps) {
        for (const precedent of deps.precedents) {
          queue.push(precedent);
        }
      }
    }
    
    visited.delete(key); // Don't include the starting node
    return visited;
  }
  
  /**
   * Detects cycles in the graph using DFS
   */
  detectCycles(): CycleDetectionResult {
    const white = new Set(this.nodes.keys()); // Not visited
    const gray = new Set<string>(); // Currently visiting
    const black = new Set<string>(); // Visited
    const parent = new Map<string, string>();
    
    const visit = (node: string): string[] | null => {
      white.delete(node);
      gray.add(node);
      
      const deps = this.dependencies.get(node);
      if (deps) {
        for (const precedent of deps.precedents) {
          if (gray.has(precedent)) {
            // Found a cycle - reconstruct the path
            const cycle = [precedent];
            let current = node;
            
            while (current !== precedent) {
              cycle.push(current);
              current = parent.get(current) || current;
            }
            
            cycle.push(precedent); // Close the cycle
            return cycle;
          }
          
          if (white.has(precedent)) {
            parent.set(precedent, node);
            const cycle = visit(precedent);
            if (cycle) return cycle;
          }
        }
      }
      
      gray.delete(node);
      black.add(node);
      return null;
    };
    
    for (const node of white) {
      const cycle = visit(node);
      if (cycle) {
        return { hasCycle: true, cycle };
      }
    }
    
    return { hasCycle: false };
  }
  
  /**
   * Find all strongly connected components using Tarjan's algorithm
   */
  findStronglyConnectedComponents(): string[][] {
    const index = new Map<string, number>();
    const lowLink = new Map<string, number>();
    const onStack = new Set<string>();
    const stack: string[] = [];
    const sccs: string[][] = [];
    let currentIndex = 0;

    const tarjan = (node: string): void => {
      index.set(node, currentIndex);
      lowLink.set(node, currentIndex);
      currentIndex++;
      stack.push(node);
      onStack.add(node);

      const deps = this.dependencies.get(node);
      if (deps) {
        for (const precedent of deps.precedents) {
          if (!index.has(precedent)) {
            // Successor not yet visited; recurse
            tarjan(precedent);
            lowLink.set(node, Math.min(lowLink.get(node)!, lowLink.get(precedent)!));
          } else if (onStack.has(precedent)) {
            // Successor is in stack and hence in the current SCC
            lowLink.set(node, Math.min(lowLink.get(node)!, index.get(precedent)!));
          }
        }
      }

      // If node is a root node, pop the stack and create an SCC
      if (lowLink.get(node) === index.get(node)) {
        const scc: string[] = [];
        let w: string;
        do {
          w = stack.pop()!;
          onStack.delete(w);
          scc.push(w);
        } while (w !== node);
        
        sccs.push(scc);
      }
    };

    for (const node of this.nodes.keys()) {
      if (!index.has(node)) {
        tarjan(node);
      }
    }

    return sccs;
  }

  /**
   * Get all nodes involved in circular dependencies using SCC analysis
   */
  getCircularNodes(): Set<string> {
    const sccs = this.findStronglyConnectedComponents();
    const circularNodes = new Set<string>();
    
    for (const scc of sccs) {
      // Only SCCs with more than one node represent cycles
      if (scc.length > 1) {
        for (const node of scc) {
          circularNodes.add(node);
        }
      }
    }
    
    return circularNodes;
  }

  /**
   * Topological sort of the graph
   */
  topologicalSort(): string[] | null {
    const inDegree = new Map<string, number>();
    const queue: string[] = [];
    const result: string[] = [];
    
    // Calculate in-degrees
    for (const [node] of this.nodes) {
      inDegree.set(node, 0);
    }
    
    for (const [, deps] of this.dependencies) {
      for (const precedent of deps.precedents) {
        inDegree.set(precedent, (inDegree.get(precedent) || 0) + 1);
      }
    }
    
    // Find nodes with no incoming edges
    for (const [node, degree] of inDegree) {
      if (degree === 0) {
        queue.push(node);
      }
    }
    
    // Process queue
    while (queue.length > 0) {
      const current = queue.shift()!;
      result.push(current);
      
      const deps = this.dependencies.get(current);
      if (deps) {
        for (const dependent of deps.dependents) {
          const degree = inDegree.get(dependent)! - 1;
          inDegree.set(dependent, degree);
          
          if (degree === 0) {
            queue.push(dependent);
          }
        }
      }
    }
    
    // Check if all nodes were processed (no cycles)
    if (result.length !== this.nodes.size) {
      return null; // Cycle detected
    }
    
    return result;
  }
  
  /**
   * Gets ranges that contain a specific cell
   */
  getRangesContainingCell(address: SimpleCellAddress): string[] {
    const cellKey = DependencyGraph.getCellKey(address);
    const ranges = new Set(this.rangeIndex.get(cellKey) || []);
    
    // Check large ranges using sparse indexing
    for (const largeRange of this.largeRanges) {
      if (this.isAddressInRange(address, largeRange.range)) {
        ranges.add(largeRange.key);
      }
    }
    
    return Array.from(ranges);
  }

  /**
   * Helper method to check if an address is within a range
   */
  private isAddressInRange(address: SimpleCellAddress, range: SimpleCellRange): boolean {
    return address.sheet === range.start.sheet &&
           address.col >= range.start.col &&
           address.col <= range.end.col &&
           address.row >= range.start.row &&
           address.row <= range.end.row;
  }
  
  /**
   * Checks if a cell is part of any tracked range
   */
  isCellInRange(address: SimpleCellAddress): boolean {
    const cellKey = DependencyGraph.getCellKey(address);
    const ranges = this.rangeIndex.get(cellKey);
    
    // Check direct index first
    if (ranges !== undefined && ranges.size > 0) {
      return true;
    }
    
    // Check large ranges
    for (const largeRange of this.largeRanges) {
      if (this.isAddressInRange(address, largeRange.range)) {
        return true;
      }
    }
    
    return false;
  }
  
  /**
   * Clear the entire graph
   */
  clear(): void {
    this.nodes.clear();
    this.dependencies.clear();
    this.rangeIndex.clear();
    this.largeRanges = [];
  }
  
  /**
   * Get the number of nodes in the graph
   */
  get size(): number {
    return this.nodes.size;
  }
  
  /**
   * Get all nodes
   */
  getAllNodes(): DependencyNode[] {
    return Array.from(this.nodes.values());
  }
  
  /**
   * Debug helper: prints the graph structure
   */
  toString(): string {
    const lines: string[] = [];
    
    for (const [key, node] of this.nodes) {
      const deps = this.dependencies.get(key)!;
      lines.push(`${key}:`);
      
      if (deps.precedents.size > 0) {
        lines.push(`  ← depends on: ${Array.from(deps.precedents).join(', ')}`);
      }
      
      if (deps.dependents.size > 0) {
        lines.push(`  → used by: ${Array.from(deps.dependents).join(', ')}`);
      }
    }
    
    return lines.join('\n');
  }
}
