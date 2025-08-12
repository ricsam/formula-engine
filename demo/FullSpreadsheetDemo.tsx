import React, { useEffect, useMemo, useState } from "react";
import { getCellReference, parseCellReference, Spreadsheet } from "@anocca-pub/components";
import { Input } from "@/components/ui/input";
import { FormulaEngine } from "../src/core/engine";
import { useSerializedSheet } from "src/react/hooks";
import type { CellAddress } from "src/core/types";

// Create a shared engine instance with rich example data
const createEngineWithExampleData = () => {
  const engine = FormulaEngine.buildEmpty();
  const sheetName = engine.addSheet("Sheet1").name;

  // Rich example data with various formulas and data types
  const exampleData = new Map<string, any>([
    // Headers
    ['A1', 'Product'],
    ['B1', 'Price'],
    ['C1', 'Quantity'],
    ['D1', 'Total'],
    ['E1', 'Tax Rate'],
    ['F1', 'Final Price'],
    
    // Data rows
    ['A2', 'Laptop'],
    ['B2', 1000],
    ['C2', 5],
    ['D2', '=B2*C2'],
    ['E2', 0.08],
    ['F2', '=D2*(1+E2)'],
    
    ['A3', 'Mouse'],
    ['B3', 30],
    ['C3', 15],
    ['D3', '=B3*C3'],
    ['E3', 0.08],
    ['F3', '=D3*(1+E3)'],
    
    ['A4', 'Keyboard'],
    ['B4', 80],
    ['C4', 8],
    ['D4', '=B4*C4'],
    ['E4', 0.08],
    ['F4', '=D4*(1+E4)'],
    
    ['A5', 'Monitor'],
    ['B5', 300],
    ['C5', 3],
    ['D5', '=B5*C5'],
    ['E5', 0.08],
    ['F5', '=D5*(1+E5)'],
    
    // Summary calculations
    ['A7', 'Summary:'],
    ['A8', 'Total Items'],
    ['B8', '=SUM(C2:C5)'],
    ['A9', 'Subtotal'],
    ['B9', '=SUM(D2:D5)'],
    ['A10', 'Total Tax'],
    ['B10', '=SUM(F2:F5)-SUM(D2:D5)'],
    ['A11', 'Grand Total'],
    ['B11', '=SUM(F2:F5)'],
    
    // Additional calculations
    ['D7', 'Analytics:'],
    ['D8', 'Avg Price'],
    ['E8', '=AVERAGE(B2:B5)'],
    ['D9', 'Max Total'],
    ['E9', '=MAX(F2:F5)'],
    ['D10', 'Min Total'],
    ['E10', '=MIN(F2:F5)'],
    
    // Text function demonstrations
    ['A13', 'Text Functions:'],
    ['A14', 'Has Laptop?'],
    ['B14', '=IF(COUNTIF(A2:A5,"Laptop")>0,"Yes","No")'],
    ['A15', 'Product List'],
    ['B15', '=CONCATENATE(A2,", ",A3,", ",A4)'],
    ['A16', 'First Product'],
    ['B16', '=UPPER(A2)'],
    ['A17', 'Name Length'],
    ['B17', '=LEN(A2)'],
    ['A18', 'Short Name'],
    ['B18', '=LEFT(A2,4)'],
    
    // INDEX function demonstrations
    ['A20', 'INDEX Examples:'],
    ['A21', '2nd Product'],
    ['B21', '=INDEX(A2:A5,2)'],
    ['A22', '3rd Price'],
    ['B22', '=INDEX(B2:B5,3)'],
    ['A23', 'Last Product'],
    ['B23', '=INDEX(A2:A5,4)'],
    ['A24', 'Dynamic Lookup'],
    ['B24', '=INDEX(A2:A5,2)&" costs $"&INDEX(B2:B5,2)'],
    
    // More text functions with INDEX
    ['A26', 'Advanced Examples:'],
    ['A27', 'Search Product'],
    ['B27', '=FIND("top",LOWER(INDEX(A2:A5,1)))'],
    ['A28', 'Replace Text'],
    ['B28', '=SUBSTITUTE(INDEX(A2:A5,1),"Lap","Note")'],
    ['A29', 'Middle Chars'],
    ['B29', '=MID(INDEX(A2:A5,1),3,3)'],
    ['A30', 'Trimmed Text'],
    ['B30', '=TRIM("  "&INDEX(A2:A5,2)&"  ")'],
    
    // Lookup table for more INDEX examples
    ['J1', 'Category'],
    ['K1', 'Description'],
    ['J2', 'Electronics'],
    ['K2', 'High-tech devices'],
    ['J3', 'Accessories'],
    ['K3', 'Supporting items'],
    ['J4', 'Computing'],
    ['K4', 'Computer equipment'],
    
    // Using INDEX with the lookup table
    ['A32', 'Category Info:'],
    ['A33', 'Cat 1 Name'],
    ['B33', '=INDEX(J2:J4,1)'],
    ['A34', 'Cat 1 Desc'],
    ['B34', '=INDEX(K2:K4,1)'],
    ['A35', 'Cat 2 Info'],
    ['B35', '=CONCATENATE(INDEX(J2:J4,2),": ",INDEX(K2:K4,2))'],
    
    // Status based on sales
    ['A37', 'Status'],
    ['A38', '=IF(B11>7000,"High Sales","Normal Sales")'],
    
    // Profit calculations
    ['H1', 'Profit Margin'],
    ['H2', '=F2*0.1'],
    ['H3', '=F3*0.1'],
    ['H4', '=F4*0.1'],
    ['H5', '=F5*0.1'],
    ['H7', 'Total Profit'],
    ['H8', '=SUM(H2:H5)'],
    
    // More INDEX with calculations
    ['H10', 'INDEX Calculations:'],
    ['H11', 'Product 1 Profit'],
    ['H12', '=INDEX(H2:H5,1)'],
    ['H13', 'Best Product'],
    ['H14', '=INDEX(A2:A5,1)&" (Best)'],
  ]);

  engine.setSheetContent(sheetName, exampleData);
  return { engine, sheetName };
};

export function FullSpreadsheetDemo() {
  const { engine, sheetName } = useMemo(createEngineWithExampleData, []);
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const spreadsheet = useSerializedSheet(engine, sheetName);
  const [formulaInput, setFormulaInput] = useState<string>("");


  console.log("Selected cell", selectedCell, spreadsheet, spreadsheet.get(selectedCell || ""));
  const formula = useMemo(() => {
    if (!selectedCell) return "";
    
    try {
      const cellFormula = spreadsheet.get(selectedCell);
      return cellFormula || "";
    } catch (error) {
      return "";
    }
  }, [engine, sheetName, selectedCell]);

  // Update formula input when selected cell changes
  useEffect(() => {
    setFormulaInput(String(formula));
  }, [formula]);

  const handleFormulaSubmit = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && selectedCell) {
      try {
        const { columnIndex, rowIndex } = parseCellReference(selectedCell);
        const address: CellAddress = { sheetName, colIndex: columnIndex, rowIndex: rowIndex };
        
        // If the input starts with =, it's a formula; otherwise it's a value
        const content = formulaInput.startsWith('=') ? formulaInput : formulaInput;
        engine.setCellContent(address, content || undefined);
      } catch (error) {
        console.error('Error updating cell:', error);
      }
    }
  };

  const handleFormulaChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFormulaInput(e.target.value);
  };

  return (
    <div className="flex flex-col gap-4 h-full w-full p-8">
      {/* Formula Bar */}
      <div className="flex gap-2 items-center">
        <div className="w-20 text-sm font-mono bg-gray-100 px-2 py-2 rounded">
          {selectedCell || ""}
        </div>
        <Input
          value={formulaInput}
          onChange={handleFormulaChange}
          onKeyDown={handleFormulaSubmit}
          placeholder={selectedCell ? "Enter formula or value..." : "Select a cell to edit"}
          className="flex-1 font-mono"
          disabled={!selectedCell}
        />
      </div>

      {/* Instructions */}
      <div className="text-sm text-muted-foreground bg-blue-50 p-3 rounded">
        <strong>Instructions:</strong>
        <ul className="list-disc list-inside mt-1 space-y-1">
          <li>Click any cell to select it and see its formula in the input above</li>
          <li>Type a formula (starting with =) or a value in the input and press Enter</li>
          <li>Try editing existing formulas like =B2*C2 or =SUM(D2:D5)</li>
          <li>Notice how dependent cells automatically update when you change values</li>
          <li>Explore the rich example data including products, calculations, and analytics</li>
        </ul>
      </div>

      {/* Spreadsheet */}
      <div className="relative flex-1">
        <Spreadsheet
          style={{ width: "100%", height: "100%" }}
          cellData={spreadsheet as Map<string, string | number>}
          onCellDataChange={(updatedSpreadsheet) => {
            engine.setSheetContent(sheetName, updatedSpreadsheet);
          }}
          customCellRenderer={(cell) => {
           
            const value = engine.getCellValue({
              sheetName,
              colIndex: cell.colIndex,
              rowIndex: cell.rowIndex,
            });
            return <div>{value}</div>;
          }}
          selection={{
            onStateChange: (state) => {
              if (state.isSelecting.type === "drag") {
                const cell = state.isSelecting.start;
                setSelectedCell(
                  getCellReference({ rowIndex: cell.row, colIndex: cell.col })
                );
                return;
              }
              const cell = state.selections[state.selections.length - 1]?.start;
              if (cell) {
                setSelectedCell(
                  getCellReference({ rowIndex: cell.row, colIndex: cell.col })
                );
              }
            },
          }}
        />
      </div>

      {/* Live Data Display */}
      <div className="bg-gray-50 p-3 rounded text-xs">
        <strong>Live Calculations:</strong>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mt-2">
          <div>Total Items: <span className="font-mono">{spreadsheet.get('B8') || 0}</span></div>
          <div>Subtotal: <span className="font-mono">${spreadsheet.get('B9') || 0}</span></div>
          <div>Grand Total: <span className="font-mono">${spreadsheet.get('B11') || 0}</span></div>
          <div>Avg Price: <span className="font-mono">${spreadsheet.get('E8') || 0}</span></div>
        </div>
      </div>
    </div>
  );
}