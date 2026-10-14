import { describe, expect, test } from "bun:test";
import { EvaluationManager } from "../../../../src/core/managers/evaluation-manager";
import type { DependencyNode } from "../../../../src/core/managers/dependency-node";

type GraphNode = DependencyNode & { edges: GraphNode[]; reads: number };
const node = (): GraphNode => {
  const value = { edges: [] as GraphNode[], reads: 0,
    getDependencies() { this.reads++; return new Set(this.edges); },
  };
  return value as unknown as GraphNode;
};
const collect = (candidates: GraphNode[], target: GraphNode) =>
  (EvaluationManager.prototype as unknown as {
    collectCircularRangeEntries(candidates: Iterable<DependencyNode>, target: DependencyNode): Set<DependencyNode>;
  }).collectCircularRangeEntries(candidates, target);

describe("range cycle reachability", () => {
  test("propagates target reachability through SCCs regardless of candidate order", () => {
    const target = node(), a = node(), b = node(), unrelated = node(), tail = node();
    a.edges = [b, target]; b.edges = [a]; unrelated.edges = [tail]; tail.edges = [unrelated];
    for (const candidates of [[a, b, unrelated], [unrelated, b, a]]) {
      const found = collect(candidates, target);
      expect(found.has(a)).toBe(true); expect(found.has(b)).toBe(true);
      expect(found.has(unrelated)).toBe(false); expect(found.has(tail)).toBe(false);
    }
    a.edges = [b];
    expect(collect([a, b], target).has(a)).toBe(false);
    b.edges.push(target);
    expect(collect([a, b], target).has(a)).toBe(true);
  });

  test("visits a shared deep dependency closure once, without recursive stack growth", () => {
    const target = node();
    const chain = Array.from({ length: 20000 }, node);
    for (let i = 1; i < chain.length; i++) chain[i]!.edges = [chain[i - 1]!];
    const candidates = Array.from({ length: 2000 }, () => {
      const candidate = node(); candidate.edges = [chain.at(-1)!]; return candidate;
    });
    expect(collect(candidates, target).size).toBe(1);
    expect(chain.every(n => n.reads === 1)).toBe(true);
    chain[0]!.edges.push(target);
    expect(collect(candidates, target).size).toBe(22001);
    expect(chain.every(n => n.reads === 2)).toBe(true);
    expect(target.reads).toBe(0);
  });

  test("matches independent DFS on deterministic random cyclic graphs", () => {
    let seed = 17;
    const random = () => (seed = (seed * 1664525 + 1013904223) >>> 0);
    for (let trial = 0; trial < 40; trial++) {
      const nodes = Array.from({ length: 50 }, node);
      for (const n of nodes) n.edges = Array.from({ length: random() % 4 }, () => nodes[random() % nodes.length]!);
      const target = nodes[0]!;
      const found = collect(nodes, target);
      for (const candidate of nodes) {
        const visited = new Set<GraphNode>();
        const stack = [candidate];
        while (stack.length) {
          const n = stack.pop()!;
          if (visited.has(n)) continue;
          visited.add(n); stack.push(...n.edges);
        }
        expect(found.has(candidate)).toBe(visited.has(target));
      }
    }
  });
});
