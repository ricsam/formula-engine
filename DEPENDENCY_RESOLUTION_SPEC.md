# Dependency Resolution Specification

## Overview

This document describes the dependency resolution system in the formula engine, with particular focus on frontier dependencies and evaluation order.

## Key Concepts

### Regular Dependencies
Dependencies that are explicitly referenced in a formula. For example, if `A1 = B1 + C1`, then A1 has regular dependencies on B1 and C1.

### Frontier Dependencies
Cells that could potentially spill values into a range referenced by a formula. These are "candidates" because we don't know if they will actually spill until they are evaluated.

### Discarded Frontier Dependencies
Frontier dependencies that have been evaluated and resolved and determined not to affect the target cell (they don't spill into the referenced range).

## Frontier Dependency Rules

### 1. Empty Cells
Empty cells can have frontier dependencies - these are cells that might spill into them.

**Example:**
```
D1 = (empty)
A1 = SEQUENCE(1, 5) // This will spill to A1:E1
```
Here, D1 would have A1 as a frontier dependency because A1 is positioned such that its spill could reach D1.

If A1 = SEQUENCE(5) then A1 spills down and D1 will remain empty. When A1 is evaluated and it is clear that A1:A5 doesn't intersect with D1 then A1 will be discarded as a frontier dependency

### 2. Cells with Formulas
Cells containing formulas can have transient dependencies with frontier dependencies - these are cells that might spill into ranges referenced by the formula.

**Example:**
```
A1 = 123
A2 = SUM(D1:D5)
B1 = SEQUENCE(3,5) // Might spill to B1:D5
```
In this case, A2 depends on D1,D2,D3,D4,D5 and open-range-evaluator would add B1 as a frontier dependency because B1 might spill into the range D1:D5 that A2 references. A1 could be a frontier dependency, however, it is blocked by B1, thus it never becomes a frontier dependency candidate.

### 3. Cells with Static Values
Cells containing static values (numbers, strings, booleans) **cannot** have frontier dependencies because they don't reference any ranges.

**Example:**
```
K5 = "hello"
```
K5 would never have frontier dependencies.

## Frontier Candidate Discovery

Frontier candidates are discovered based on their position relative to a range:

1. **Above the range**: Cells in the same columns but in rows before the range
2. **To the left of the range**: Cells in the same rows but in columns before the range
3. **Top-left quadrant**: Cells both above AND to the left of the range

The discovery process checks for "blocking" cells - non-empty cells between the candidate and the range that would prevent spilling. E.g. if the range starts with a formula e.g. SEQUENCE(5), then that cell with the formula will block any candidates above. If the range is 3 columns and the first column is filled with static values, then no frontier dependencies to the left can be found.

## Evaluation Order

The evaluation order for a cell is determined by **dependency-based topological sorting** rather than artificial phase priorities. The algorithm ensures:

1. **Dependencies come before dependents** - If cell A depends on cell B, then B is evaluated before A
2. **Frontier dependencies are treated as soft dependencies** - They don't create cycles but are included in the dependency graph
3. **Target cell is included** - The cell being evaluated appears in its own evaluation order
4. **Iterative discovery** - New dependencies discovered during evaluation trigger re-ordering

**Key principle**: The evaluation order respects the actual dependency relationships in the spreadsheet, ensuring correct results while avoiding artificial constraints.

### Example Evaluation Order with SUM

Given:
```
A1 = 1
A2 = SEQUENCE(F1, 2)
A3 = 3
B1 = D11 * 0.5
B2 = ""
B3 = 7
C1 = SUM(A1:A3 * B1:B3)
D10 = A1:A2 * (B2 + A1) // Spills to D10:D11
F1 = 1
```

When evaluating C1, the engine follows an iterative dependency discovery process:

#### Step 1: Initial Discovery
- **Dependencies discovered**: B1, A1 (from multiplication of first elements)
- **Frontier dependencies discovered**: A2 (frontier candidate for B1:B3 range)
- **Evaluation order**: `["B1", "A1", "A2", "C1"]`

#### Step 2: Transitive Discovery
- **New dependencies**: F1 (from A2), D11 (from B1)
- **Evaluation order**: `["D11", "B1", "A1", "F1", "A2", "C1"]`

#### Step 3: Frontier Discovery for Empty Cells
- **New frontier dependencies**: D10 (frontier candidate for empty D11)
- **Evaluation order**: `["D10", "F1", "A2", "D11", "B1", "A1", "C1"]`

#### Step 4: Spill Resolution
- D10 spills to D11, creating regular dependencies A1, A2, B2 for D11
- **New dependencies**: A1, A2, B2 (for D11), B3, A3 (for C1 continuation)
- **Evaluation order**: `["F1", "A2", "A1", "D10", "D11", "B1", "B3", "A3", "C1"]`

#### Final Step: Complete Evaluation
- All dependencies resolved, C1 evaluates successfully
- **Final evaluation order**: `["F1", "A2", "A1", "D10", "D11", "B1", "B3", "A3", "C1"]`

**Key observations:**
1. **Iterative discovery**: Dependencies are discovered progressively over multiple evaluation runs
2. **Dependency-based ordering**: The evaluation order respects dependency relationships (F1 before A2, D10 before D11, etc.)
3. **Frontier upgrade**: D10 starts as a frontier dependency of D11 but becomes a regular dependency when it spills
4. **No artificial phases**: The algorithm doesn't enforce "frontier before regular" but ensures correct dependency resolution

Performance optimizations not mentioned:
 * a frontier dependency will not be added if it is already a dependency
 * cells marked as "resolved" (i.e. after re-evaluation of a cell no additional nodes nor frontier dependencies were discovered) which do not spill or spill without intersecting should not become frontier dependencies. For example when A2 is added as a frontier dependency to D11, then if A2 was already "resolved", i.e. F1 is evaluated then D11 doesn't become a frontier dependency
 * the evaluation order list should not include duplicates
 * when empty cells are evaluated the EvaluationManager will fist check if any cell spills into the cell, if not then frontier dependencies will be added so that they can be checked
 * when the OpenRangeEvaluator evaluates open ranges and it fins that a frontier dependency spills into the range, then that frontier dependency will also be upgraded to a real dependency

### Dependency types
 * normal dependencies of A are added when the forumla in A references another cells. Upgraded dependencies, i.e. when a frontier dependency is "upgraded" to a normal dependency, will also be added to the same list of normal dependencies.
 * frontier dependencies of A are other formula cells that could pontentially influence A when spilling
 * discarded frontier dependencies of A are resolved formula cells which never spilled into the A

### Dependency upgrades/downgrade and discovery
 * The OpenRangeEvaluator discovers frontier dependency to ranges it evaluates, it also disovers normal dependencies residing within the range
 * EvaluationManager discoveres frontier dependencies when evaluating emtpy cells
 * EvaluationManager discovers normal dependencies when evaluating formulas, and ranges are passed to the OpenRangeEvaluator which will pass back the found dependencies / frontier dependencies to the EvaluationManager.
 * A frontier dependency can only be upgraded to a true dependency if the dependeny is marked as resolved, because a cell could have a B1=IFERROR(A1, SEQUENCE(10,10) and thus falsly spill before A1 is resolved
 * Both the OpenRangeEvaluator and the EvaluationManager can upgrade/discard frontier dependencies
 * The OpenRangeEvaluator upgrades resolved frontier dependencies that spill into the evaluated range, or downgrade resolved spills that do not intersect
 * The EvaluationManager will upgrade/add a dependency/downgrade when it evaluates an empty cell and it is within the spill range of some resolved cell or downgrade the dependency
 * When a cell spills the EvaluationManager will NOT upgrade or add any dependencies, it will simply register the spill. Spilling cells should according to the algorithm always evaluate before cells that potentially consume them

### Example Evaluation Order with Multiplication

Given:
```
A1 = 1
A2 = SEQUENCE(F1, 2)
A3 = 3
B1 = D11 * 0.5
B2 = ""
B3 = 7
C1 = A1:A3 * B1:B3 // Array multiplication (no SUM)
D10 = A1:A2 * (B2 + A1) // Spills to D10:D11
F1 = 1
```

When evaluating C1:
1. **Initial discovery**: A1, B1 (multiplication evaluates first elements)
2. **Eval order**: `["A1", "B1", "C1"]`
3. **New discovery**: D11 (from B1 evaluation)
4. **Eval order**: `["D11", "A1", "B1", "C1"]`
5. **Frontier discovery**: D10 (frontier candidate for empty D11)
6. **Eval order**: `["D10", "D11", "A1", "B1", "C1"]`
7. **Spill resolution**: D10 spills to D11, creates dependencies A1, A2, B2 for D11
8. **Final eval order**: All transitive dependencies resolved, C1 evaluates correctly

**Key difference from SUM**: Multiplication doesn't use `evaluateAllCells` so it discovers dependencies more gradually, but follows the same dependency-based ordering principles.

### Example with Nested Array Dependencies

Given:
```
A1 = SUM(C3:D3)
B2 = I12:K14          // Array formula referencing range I12:K14
H10 = SEQUENCE(10, 10) // Spills to H10:Q19
```

When evaluating A1:
1. Initial discovery: [B2] (frontier dep for empty C3:D3)
2. Eval order: [B2], A1
3. New discovery: [H10] (frontier dep for empty I12:K14 range referenced by B2)
4. Eval order: [H10], [B2], A1
5. H10 spills to H10:Q19, intersecting I12:K14, upgraded to {H10} dependency of B2
6. Eval order: [H10], {H10}, B2, A1
7. B2 evaluates with spilled values from I12:K14, spills to B2:D4, upgraded to {B2} dependency of A1
8. Final eval order: H10, B2, A1

## Cycle Detection

Cycles are detected only among **regular dependencies**. The algorithm uses a specialized topological sort that:

1. **Detects cycles in regular dependencies** - Uses standard cycle detection on the regular dependency graph
2. **Treats frontier dependencies as soft edges** - Includes them in evaluation order but doesn't consider them for cycle detection
3. **Prevents infinite loops** - Frontier dependencies can't create visiting cycles during topological traversal

### Example of Non-Cycle with Frontier Dependencies

```
C1 = SUM(A1:A3)
B1 = D11 * 0.5
D10 = C1 + 5 // Depends on C1
```

This scenario does **not** create a cycle even though there's a dependency chain C1 → (frontier) → B1 → D11 → D10 → C1 because:
- **B1 is a frontier dependency** of C1 (soft edge)
- **D10 → C1 is a regular dependency** (hard edge)  
- **Only regular dependencies are considered for cycle detection**
- The evaluation order algorithm handles this correctly: `[D10, D11, B1, A1, A3, C1]`

## Implementation Details

### Dependency Discovery
- Regular dependencies are discovered during formula parsing
- Frontier dependencies are discovered based on cell position and spill potential
- Empty cells discover frontier candidates when they are evaluated
- Formulas discover frontier candidates for ranges they reference

### Dependency Storage
- Each evaluated node stores:
  - `deps`: Set of regular dependencies (node keys)
  - `frontierDependencies`: Map<string, Set<string>> - Range-keyed active frontier dependencies
  - `discardedFrontierDependencies`: Map<string, Set<string>> - Range-keyed discarded frontier dependencies

**Range-keyed structure example:**
```typescript
{
  deps: Set(["cell:Workbook:Sheet:A1"]),
  frontierDependencies: Map([
    ["B1:B3", Set(["cell:Workbook:Sheet:A2"])],
    ["D1:D5", Set(["cell:Workbook:Sheet:C1"])]
  ]),
  discardedFrontierDependencies: Map([
    ["B1:B3", Set(["cell:Workbook:Sheet:A3"])]
  ])
}
```

### Performance Optimization
- **Frontier dependency indexes** are maintained to avoid O(n) lookups in StoreManager:
  - `frontierDependencyIndex`: Maps frontier dep → cells that reference it
  - `discardedFrontierDependencyIndex`: Maps discarded frontier dep → cells that referenced it

- **Evaluation order hashing**: The `buildEvaluationOrder()` method returns a hash representing the current state of all evaluated nodes, including:
  - Regular dependencies
  - Frontier dependencies (by range)
  - Discarded frontier dependencies (by range)
  
  This hash changes when any dependency discovery or discard occurs, enabling efficient re-evaluation detection.

## Edge Cases

### Re-evaluation Triggers
A cell needs re-evaluation when:
1. Its regular dependencies change value
2. A frontier dependency starts or stops spilling into its range
3. New dependencies are discovered during evaluation

### Overlapping spills
A1 = SEQUENCE(2,3)  // Spills to A1:C2
B1 = SEQUENCE(2,3)  // Spills to B1:D2

This would cause A1 to have a SPILL error. B1 would evaluate fine.

### Chained spills
A1 = SEQUENCE(2,2)     // Spills to A1:B2
C1 = A1 * 2            // Also spills to C1:D2
E1 = SUM(F1:G2)        // Depends on range that C1 might spill into

evaluateCellsInRange runs against F1:G2. If that range is empty no dependencies are added.
F1 would have one frontier dependency: E1 which later would be discarded as E1 doesn't intersect with F1:G2.
Note that it is the source range that should be checked for fronier dependencies, not E1:F2 (and this is how it is implemented when looking at evaluateRange in FormulaEvaluator and the OpenRangeEvaluator), however, the context in which
the range is evaluated is E1 meaning that the frontier dependencies are added to E1. Therefore E1 could have itself as a frontier dependency (however, right now there is logic in the OpenRangeEvaluator to skip adding frontier dependencies that are the same as the context origin, but that logic is subject to change to make the code complient with the algorithm presented in this document)

### Frontier Dependencies Through Indirect References
INDERECT has not been implemented yet. But during evaluation the reference is resolved just like it does for direct references in this engine.

## Testing Considerations

When testing dependency resolution:
1. Test both static and dynamic dependency discovery
2. Verify frontier candidates are correctly identified for all three regions
3. Ensure evaluation order respects the priority rules
4. Confirm cycles are only detected for regular dependencies
5. Test spill restoration when frontier dependencies change their spill behavior (e.g. before all dependencies are evaluated a spilling cell may just evaluate to Error, but later once all dependencies are evaluated it may actually spill)


## Concerned files
1. open-range-evaluator.ts - to get all cells in an open ended range and add frontier dependencies
2. workbook-manager.ts - methods for doing the frontier cell lookups
3. store-manager.ts - methods for adding and removing dependencies on a cell
4. evaluation-manager.ts - conntains the main evaluation loop
5. dependency-manager.ts - calculates the evaluation order

## Dependency Lifecycle
This is subject to change, but as it stands right now:

### Discovered
All dependencies are added to a cell in the EvaluationManager.evaluateDependencyNode.
When a cell formula is evaluated dependencies are attached to the following sets:
* dependenciesDiscoveredInEvaluation
* frontierDependenciesDiscoveredInEvaluation
* discardedFrontierDependenciesDiscoveredInEvaluation

When a formula references another cell then that cell will be added to the dependencies.

When evaluateAllCells is called on a range, e.g for SUM, or AVERAGE then the OpenRangeEvaluator.evaluateCellsInRange is
called. This method will add and remove frontier dependencies and dependencies. This is absolutely subject to change and I know the current implementation isn't correct given the examples above.

Aditionally, in evaluateDependencyNode, if an evaluated cell is just an empty string then frontier dependencies are looked up. And if a cell evaluates and starts spilling it will restore itself as a frontier dependency for intersecting
cells that previously discoarded it.

### Evaluation Order Algorithm
The `DependencyManager.buildEvaluationOrder()` method implements a **specialized topological sort** that:

1. **Collects all nodes** in the dependency graph (regular + frontier dependencies)
2. **Separates node types** into frontier and regular dependency categories  
3. **Checks for cycles** in regular dependencies only (frontier dependencies don't create cycles)
4. **Performs specialized DFS traversal**:
   - Visit regular dependencies first (hard edges)
   - Visit frontier dependencies second (soft edges, but avoid visiting cycles)
   - Ensures dependencies come before dependents

The `EvaluationManager.evaluateCell()` continues in a while loop until no more dependencies are discovered, using the computed evaluation order to call `evaluateDependencyNode()` in the correct sequence.

