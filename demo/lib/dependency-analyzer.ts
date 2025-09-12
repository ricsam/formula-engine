import type { SerializedCellValue } from "src/core/types";
import { FormulaEngine } from "../../src/core/engine";

// Helper functions for column conversion
function columnToIndex(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result - 1; // Convert to 0-based
}

function indexToColumn(index: number): string {
  let result = '';
  let num = index + 1; // Convert to 1-based
  while (num > 0) {
    num--;
    result = String.fromCharCode('A'.charCodeAt(0) + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

export interface SheetDependency {
  fromSheet: string;
  toSheet: string;
  formulas: Array<{ formula: SerializedCellValue; cellAddress: string; }>;
  cellReferences: string[];
}

export interface DependencyGraph {
  nodes: Array<{ id: string; name: string; emoji: string }>;
  edges: Array<{ 
    id: string; 
    source: string; 
    target: string; 
    formulas: Array<{ formula: SerializedCellValue; cellAddress: string; }>;
    cellCount: number;
  }>;
}

/**
 * Extracts cross-sheet references from a formula string
 * Examples: "Products!A1", "Sales!B2:C10", "Dashboard!A1:B5"
 */
function extractCrossSheetReferences(formula: SerializedCellValue): Array<{ sheet: string; range: string }> {
  // Match patterns like SheetName!CellRange
  const crossSheetPattern = /([A-Za-z_][A-Za-z0-9_]*)\!([A-Z]+\d+(?:\:[A-Z]+\d+)?)/g;
  const references: Array<{ sheet: string; range: string }> = [];

  if (typeof formula !== 'string') {
    return [];
  }
  
  let match;
  while ((match = crossSheetPattern.exec(formula)) !== null) {
    if (match[1] && match[2]) {
      references.push({
        sheet: match[1],
        range: match[2],
      });
    }
  }
  
  return references;
}

/**
 * Analyzes the formula engine to extract all cross-sheet dependencies
 */
export function analyzeDependencies(
  engine: FormulaEngine,
  sheets: { [key: string]: { name: string; workbookName: string; } }
): DependencyGraph {
  const dependencies = new Map<string, SheetDependency>();
  
  console.log('Starting dependency analysis...');
  
  // Iterate through all sheets and their cells
  Object.entries(sheets).forEach(([sheetKey, sheet]) => {
      const sheetFormulas = engine.getSheetSerialized({ sheetName: sheet.name, workbookName: sheet.workbookName });
    console.log(`Analyzing sheet: ${sheet.name} (${sheetKey}) with ${sheetFormulas.size} formulas`);
    
    let crossSheetCount = 0;
    
    sheetFormulas.forEach((formula: SerializedCellValue, cellAddress: string) => {
      console.log(`Found formula in ${sheet.name}!${cellAddress}: ${formula}`);
      
      const crossSheetRefs = extractCrossSheetReferences(formula);
      
      if (crossSheetRefs.length > 0) {
        crossSheetCount++;
        console.log(`Cross-sheet refs found:`, crossSheetRefs);
      }
      
      crossSheetRefs.forEach(ref => {
        const dependencyKey = `${sheet.name}->${ref.sheet}`;
        
        if (!dependencies.has(dependencyKey)) {
          dependencies.set(dependencyKey, {
            fromSheet: sheet.name,
            toSheet: ref.sheet,
            formulas: [],
            cellReferences: [],
          });
        }
        
        const dependency = dependencies.get(dependencyKey)!;
        
        // Add formula with cell address if not already present
        if (!dependency.formulas.some(f => f.formula === formula && f.cellAddress === cellAddress)) {
          dependency.formulas.push({ formula, cellAddress });
        }
        
        // Add cell reference
        dependency.cellReferences.push(`${cellAddress}: ${ref.sheet}!${ref.range}`);
      });
    });
    
    console.log(`Sheet ${sheet.name}: ${sheetFormulas.size} formulas, ${crossSheetCount} cross-sheet references`);
  });

  // Create nodes for all sheets
  const nodes = Object.values(sheets).map(sheet => ({
    id: sheet.name,
    name: sheet.name,
    emoji: getSheetEmoji(sheet.name),
  }));

  // Create edges for dependencies
  const edges = Array.from(dependencies.values()).map((dep, index) => ({
    id: `edge-${index}`,
    source: dep.toSheet, // Data flows FROM the referenced sheet
    target: dep.fromSheet, // TO the sheet with the formula
    formulas: dep.formulas, // Show all formulas
    cellCount: dep.cellReferences.length,
  }));

  console.log(`Total dependencies found: ${dependencies.size}`);
  console.log('Dependencies:', Array.from(dependencies.entries()));
  console.log('Final edges:', edges);

  return { nodes, edges };
}

function getSheetEmoji(sheetName: string): string {
  switch (sheetName.toLowerCase()) {
    case 'products':
      return '📦';
    case 'sales':
      return '💰';
    case 'dashboard':
      return '📊';
    default:
      return '📄';
  }
}

/**
 * Simplifies formulas for display by removing long ranges and focusing on the key parts
 */
export function simplifyFormulaForDisplay(formula: string): string {
  return formula
    .replace(/\!([A-Z]+)\d+\:\1\d+/g, '!$1:$1') // A2:A10000 -> A:A
    .replace(/\d{4,}/g, '...') // Replace long numbers with ...
    .substring(0, 120) + (formula.length > 120 ? '...' : ''); // Longer truncate for full view
}

/**
 * Enhanced formula display that shows more context without simplification
 */
export function enhancedFormulaForDisplay(formula: string, cellAddress?: string): string {
  // Show the full formula but with better formatting
  let display = formula;
  
  // Add line breaks for very long formulas at logical points
  display = display.replace(/,(?=[A-Z])/g, ',\n  '); // Break after commas before new arguments
  display = display.replace(/\)\s*\+\s*/g, ') +\n  '); // Break after function calls when adding
  
  return display;
}
