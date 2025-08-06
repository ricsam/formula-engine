import {
  getCellReference,
  parseCellReference,
  Spreadsheet,
} from "@anocca-pub/components";
import React, { useEffect, useMemo, useState, useCallback } from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { FormulaEngine } from "../src/core/engine";

const createEngineWithMultiSheetData = () => {
  const engine = FormulaEngine.buildEmpty();

  // Create three sheets
  const salesSheet = engine.addSheet("Sales");
  const productsSheet = engine.addSheet("Products");
  const dashboardSheet = engine.addSheet("Dashboard");

  const salesSheetId = engine.getSheetId(salesSheet);
  const productsSheetId = engine.getSheetId(productsSheet);
  const dashboardSheetId = engine.getSheetId(dashboardSheet);

  // Products Sheet - Master product data
  const productsData = new Map<string, any>([
    // Headers
    ["A1", "Product ID"],
    ["B1", "Product Name"],
    ["C1", "Category"],
    ["D1", "Unit Price"],
    ["E1", "Cost"],
    ["F1", "Margin"],

    // Product data
    ["A2", "P001"],
    ["B2", "Gaming Laptop"],
    ["C2", "Electronics"],
    ["D2", 1200],
    ["E2", 800],
    ["F2", "=(D2-E2)/D2"],

    ["A3", "P002"],
    ["B3", "Wireless Mouse"],
    ["C3", "Accessories"],
    ["D3", 45],
    ["E3", 20],
    ["F3", "=(D3-E3)/D3"],

    ["A4", "P003"],
    ["B4", "Mechanical Keyboard"],
    ["C4", "Accessories"],
    ["D4", 120],
    ["E4", 60],
    ["F4", "=(D4-E4)/D4"],

    ["A5", "P004"],
    ["B5", "4K Monitor"],
    ["C5", "Electronics"],
    ["D5", 350],
    ["E5", 200],
    ["F5", "=(D5-E5)/D5"],

    ["A6", "P005"],
    ["B6", "Tablet"],
    ["C6", "Electronics"],
    ["D6", 600],
    ["E6", 400],
    ["F6", "=(D6-E6)/D6"],

    // Summary calculations
    ["A8", "Summary:"],
    ["A9", "Total Products"],
    ["B9", "=COUNT(A2:A6)"],
    ["A10", "Avg Price"],
    ["B10", "=AVERAGE(D2:D6)"],
    ["A11", "Avg Margin"],
    ["B11", "=AVERAGE(F2:F6)"],

    // Category analysis
    ["D8", "By Category:"],
    ["D9", "Electronics Count"],
    ["E9", '=COUNTIF(C2:C6,"Electronics")'],
    ["D10", "Accessories Count"],
    ["E10", '=COUNTIF(C2:C6,"Accessories")'],
  ]);

  // Sales Sheet - Transaction data with cross-sheet references
  const salesData = new Map<string, any>([
    // Headers
    ["A1", "Sale ID"],
    ["B1", "Date"],
    ["C1", "Product ID"],
    ["D1", "Product Name"],
    ["E1", "Quantity"],
    ["F1", "Unit Price"],
    ["G1", "Total"],
    ["H1", "Category"],

    // Sales transactions
    ["A2", "S001"],
    ["B2", "2024-01-15"],
    ["C2", "P001"],
    ["D2", "=INDEX(Products.B:B,MATCH(C2,Products.A:A,0))"],
    ["E2", 2],
    ["F2", "=INDEX(Products.D:D,MATCH(C2,Products.A:A,0))"],
    ["G2", "=E2*F2"],
    ["H2", "=INDEX(Products.C:C,MATCH(C2,Products.A:A,0))"],

    ["A3", "S002"],
    ["B3", "2024-01-16"],
    ["C3", "P002"],
    ["D3", "=INDEX(Products.B:B,MATCH(C3,Products.A:A,0))"],
    ["E3", 5],
    ["F3", "=INDEX(Products.D:D,MATCH(C3,Products.A:A,0))"],
    ["G3", "=E3*F3"],
    ["H3", "=INDEX(Products.C:C,MATCH(C3,Products.A:A,0))"],

    ["A4", "S003"],
    ["B4", "2024-01-17"],
    ["C4", "P003"],
    ["D4", "=INDEX(Products.B:B,MATCH(C4,Products.A:A,0))"],
    ["E4", 3],
    ["F4", "=INDEX(Products.D:D,MATCH(C4,Products.A:A,0))"],
    ["G4", "=E4*F4"],
    ["H4", "=INDEX(Products.C:C,MATCH(C4,Products.A:A,0))"],

    ["A5", "S004"],
    ["B5", "2024-01-18"],
    ["C5", "P004"],
    ["D5", "=INDEX(Products.B:B,MATCH(C5,Products.A:A,0))"],
    ["E5", 1],
    ["F5", "=INDEX(Products.D:D,MATCH(C5,Products.A:A,0))"],
    ["G5", "=E5*F5"],
    ["H5", "=INDEX(Products.C:C,MATCH(C5,Products.A:A,0))"],

    ["A6", "S005"],
    ["B6", "2024-01-19"],
    ["C6", "P005"],
    ["D6", "=INDEX(Products.B:B,MATCH(C6,Products.A:A,0))"],
    ["E6", 2],
    ["F6", "=INDEX(Products.D:D,MATCH(C6,Products.A:A,0))"],
    ["G6", "=E6*F6"],
    ["H6", "=INDEX(Products.C:C,MATCH(C6,Products.A:A,0))"],

    ["A7", "S006"],
    ["B7", "2024-01-20"],
    ["C7", "P001"],
    ["D7", "=INDEX(Products.B:B,MATCH(C7,Products.A:A,0))"],
    ["E7", 1],
    ["F7", "=INDEX(Products.D:D,MATCH(C7,Products.A:A,0))"],
    ["G7", "=E7*F7"],
    ["H7", "=INDEX(Products.C:C,MATCH(C7,Products.A:A,0))"],

    // Sales summary
    ["A9", "Sales Summary:"],
    ["A10", "Total Sales"],
    ["B10", "=SUM(G2:G7)"],
    ["A11", "Total Units"],
    ["B11", "=SUM(E2:E7)"],
    ["A12", "Avg Sale Value"],
    ["B12", "=AVERAGE(G2:G7)"],

    // Category breakdown
    ["D9", "By Category:"],
    ["D10", "Electronics Sales"],
    ["E10", '=SUMIF(H2:H7,"Electronics",G2:G7)'],
    ["D11", "Accessories Sales"],
    ["E11", '=SUMIF(H2:H7,"Accessories",G2:G7)'],

    // Product performance
    ["D13", "Top Products:"],
    ["D14", "P001 Sales"],
    ["E14", '=SUMIF(C2:C7,"P001",G2:G7)'],
    ["D15", "P002 Sales"],
    ["E15", '=SUMIF(C2:C7,"P002",G2:G7)'],
  ]);

  // Dashboard Sheet - Aggregated data from both sheets
  const dashboardData = new Map<string, any>([
    ["A1", "BUSINESS DASHBOARD"],

    // Product overview
    ["A3", "PRODUCT OVERVIEW"],
    ["A4", "Total Products"],
    ["B4", "=Products.B9"],
    ["A5", "Average Price"],
    ["B5", "=Products.B10"],
    ["A6", "Average Margin"],
    ["B6", '=CONCATENATE(ROUND(Products.B11*100,1),"%")'],

    // Sales overview
    ["A8", "SALES OVERVIEW"],
    ["A9", "Total Revenue"],
    ["B9", "=Sales.B10"],
    ["A10", "Total Units Sold"],
    ["B10", "=Sales.B11"],
    ["A11", "Average Sale Value"],
    ["B11", "=Sales.B12"],

    // Category performance
    ["D3", "CATEGORY PERFORMANCE"],
    ["D4", "Electronics"],
    ["E4", "Products:"],
    ["F4", "=Products.E9"],
    ["G4", "Sales:"],
    ["H4", "=Sales.E10"],
    ["D5", "Accessories"],
    ["E5", "Products:"],
    ["F5", "=Products.E10"],
    ["G5", "Sales:"],
    ["H5", "=Sales.E11"],

    // Performance metrics
    ["A13", "PERFORMANCE METRICS"],
    ["A14", "Revenue per Product"],
    ["B14", "=Sales.B10/Products.B9"],
    ["A15", "Conversion Rate"],
    ["B15", '=CONCATENATE(ROUND((Sales.B11/Products.B9)*100,1),"%")'],

    // Top performing products
    ["D7", "TOP PRODUCTS"],
    ["D8", "Gaming Laptop Sales"],
    ["E8", "=Sales.E14"],
    ["D9", "Wireless Mouse Sales"],
    ["E9", "=Sales.E15"],

    // Dynamic lookups
    ["A17", "PRODUCT LOOKUP"],
    ["A18", "Product ID:"],
    ["B18", "P001"],
    ["A19", "Product Name:"],
    ["B19", "=INDEX(Products.B:B,MATCH(B18,Products.A:A,0))"],
    ["A20", "Category:"],
    ["B20", "=INDEX(Products.C:C,MATCH(B18,Products.A:A,0))"],
    ["A21", "Unit Price:"],
    ["B21", "=INDEX(Products.D:D,MATCH(B18,Products.A:A,0))"],
    ["A22", "Total Sales:"],
    ["B22", "=SUMIF(Sales.C:C,B18,Sales.G:G)"],

    // Text functions showcase
    ["D11", "TEXT ANALYSIS"],
    ["D12", "Best Category"],
    ["E12", '=IF(Sales.E10>Sales.E11,"Electronics","Accessories")'],
    ["D13", "Report Title"],
    ["E13", '=CONCATENATE("Sales Report - ",UPPER(E12)," LEADING")'],
    ["D14", "Summary"],
    ["E14", '=CONCATENATE("Total: $",Sales.B10," from ",Sales.B11," units")'],
  ]);

  // Populate all sheets
  engine.setSheetContents(productsSheetId, productsData);
  engine.setSheetContents(salesSheetId, salesData);
  engine.setSheetContents(dashboardSheetId, dashboardData);

  return {
    engine,
    sheets: {
      sales: { name: salesSheet, id: salesSheetId },
      products: { name: productsSheet, id: productsSheetId },
      dashboard: { name: dashboardSheet, id: dashboardSheetId },
    },
  };
};

export function MultiSheetDemo() {
  const { engine, sheets } = useMemo(createEngineWithMultiSheetData, []);
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [activeSheet, setActiveSheet] = useState<string>("Dashboard");
  const [formulaInput, setFormulaInput] = useState<string>("");
  const [spreadsheets, setSpreadsheets] = useState<
    Record<string, Map<string, any>>
  >(() => ({
    Dashboard: engine.getSheetSerialized(sheets.dashboard.id),
    Sales: engine.getSheetSerialized(sheets.sales.id),
    Products: engine.getSheetSerialized(sheets.products.id),
  }));

  // Update all spreadsheets when any cell changes
  useEffect(() => {
    const unsubscribe = engine.on("cell-changed", () => {
      setSpreadsheets({
        Dashboard: engine.getSheetSerialized(sheets.dashboard.id),
        Sales: engine.getSheetSerialized(sheets.sales.id),
        Products: engine.getSheetSerialized(sheets.products.id),
      });
    });
    return () => unsubscribe();
  }, [engine, sheets]);

  const activeSheetId = useMemo(() => {
    const sheetEntry = Object.values(sheets).find(
      (sheet) => sheet.name === activeSheet
    );
    return sheetEntry?.id ?? sheets.dashboard.id;
  }, [activeSheet, sheets]);

  const formula = useMemo(() => {
    if (!selectedCell) return "";
    try {
      const { columnIndex, rowIndex } = parseCellReference(selectedCell);
      const cellFormula = engine.getCellFormula({
        sheet: activeSheetId,
        col: columnIndex,
        row: rowIndex,
      });
      return cellFormula || "";
    } catch (error) {
      return "";
    }
  }, [activeSheetId, selectedCell, engine]);

  // Update formula input when selected cell changes
  useEffect(() => {
    setFormulaInput(formula);
  }, [formula]);

  const handleFormulaSubmit = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === 'Enter' && selectedCell) {
        try {
          const { columnIndex, rowIndex } = parseCellReference(selectedCell);
          const address = { sheet: activeSheetId, col: columnIndex, row: rowIndex };
          
          // If the input starts with =, it's a formula; otherwise it's a value
          const content = formulaInput.startsWith('=') ? formulaInput : formulaInput;
          engine.setCellContents(address, content || undefined);
        } catch (error) {
          console.error('Error updating cell:', error);
        }
      }
    },
    [selectedCell, activeSheetId, engine, formulaInput]
  );

  const handleFormulaChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      setFormulaInput(e.target.value);
    },
    []
  );

  const addNewSale = useCallback(() => {
    const salesData = engine.getSheetContents(sheets.sales.id);
    const salesKeys = Array.from(salesData.keys()).filter(
      (key) => key.startsWith("A") && key !== "A1"
    );
    const lastRow = Math.max(
      ...salesKeys.map((key) => parseInt(key.substring(1)))
    );
    const newRow = lastRow + 1;

    // Add a new sale with random product
    const productIds = ["P001", "P002", "P003", "P004", "P005"];
    const randomProductId =
      productIds[Math.floor(Math.random() * productIds.length)];
    const randomQuantity = Math.floor(Math.random() * 5) + 1;

    const newSaleData = new Map<string, any>([
      [`A${newRow}`, `S${String(newRow - 1).padStart(3, "0")}`],
      [`B${newRow}`, "2024-01-21"],
      [`C${newRow}`, randomProductId],
      [`D${newRow}`, `=INDEX(Products.B:B,MATCH(C${newRow},Products.A:A,0))`],
      [`E${newRow}`, randomQuantity],
      [`F${newRow}`, `=INDEX(Products.D:D,MATCH(C${newRow},Products.A:A,0))`],
      [`G${newRow}`, `=E${newRow}*F${newRow}`],
      [`H${newRow}`, `=INDEX(Products.C:C,MATCH(C${newRow},Products.A:A,0))`],
    ]);

    newSaleData.forEach((value, key) => {
      const { columnIndex, rowIndex } = parseCellReference(key);
      engine.setCellContents(
        { sheet: sheets.sales.id, col: columnIndex, row: rowIndex },
        value
      );
    });
  }, [engine, sheets.sales.id]);

  const createSpreadsheetComponent = (sheetName: string, sheetId: number) => (
    <div className="flex flex-col gap-2">
      <div className="flex items-center justify-between">
        <h3 className="text-lg font-semibold">
          {sheetName === "Dashboard" && "📊 Dashboard"}
          {sheetName === "Sales" && "💰 Sales"}
          {sheetName === "Products" && "📦 Products"}
        </h3>
        <div className="text-xs text-gray-500">
          {activeSheet === sheetName && selectedCell && (
            <span>Selected: {selectedCell}</span>
          )}
        </div>
      </div>
      <div className="border rounded-lg overflow-hidden bg-white" style={{ height: "400px" }}>
        <Spreadsheet
          style={{ width: "100%", height: "100%" }}
          cellData={spreadsheets[sheetName]}
          onCellDataChange={(updatedSpreadsheet) => {
            engine.setSheetContents(sheetId, updatedSpreadsheet);
          }}
          customCellRenderer={(cell) => {
            const value = engine.getCellValue({
              sheet: sheetId,
              col: cell.colIndex,
              row: cell.rowIndex,
            });
            return <div>{value}</div>;
          }}
          selection={{
            onStateChange: (state) => {
              // Set this sheet as active when user interacts with it
              setActiveSheet(sheetName);
              
              if (state.isSelecting?.type === "drag") {
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
    </div>
  );

  return (
    <div className="flex flex-col gap-4 h-full">
      {/* Header with controls */}
      <div className="flex items-center gap-4 p-4 border-b">
        <h2 className="text-xl font-bold">Multi-Sheet Demo - Cross-Sheet References</h2>
        <Button onClick={addNewSale} variant="outline">
          Add Random Sale
        </Button>
      </div>

      {/* Formula bar */}
      <div className="flex items-center gap-2 px-4">
        <span className="text-sm font-medium min-w-20">
          {selectedCell ? `${activeSheet}!${selectedCell}` : "No cell selected"}
        </span>
        <Input
          value={formulaInput}
          onChange={handleFormulaChange}
          onKeyDown={handleFormulaSubmit}
          className="flex-1 font-mono"
          placeholder={selectedCell ? "Enter formula or value..." : "Select a cell to edit"}
          disabled={!selectedCell}
        />
      </div>

      {/* Three spreadsheets side by side */}
      <div className="flex-1 overflow-hidden px-4">
        <div className="grid grid-cols-3 gap-4 h-full">
          {createSpreadsheetComponent("Products", sheets.products.id)}
          {createSpreadsheetComponent("Sales", sheets.sales.id)}
          {createSpreadsheetComponent("Dashboard", sheets.dashboard.id)}
        </div>
      </div>

      {/* Key features info */}
      <div className="px-4 py-2 bg-gray-50 border-t text-xs">
        <strong>Key Features Demonstrated:</strong>
        <span className="ml-2">
          Cross-sheet references (Products.A1, Sales.B10) • INDEX/MATCH lookups
          • SUMIF/COUNTIF aggregations • Text functions (CONCATENATE, UPPER) •
          Dynamic calculations • Real-time updates across sheets
        </span>
      </div>

      {/* Live summary */}
      <div className="px-4 py-2 bg-blue-50 text-xs">
        <strong>Live Summary:</strong>
        <div className="grid grid-cols-3 gap-4 mt-1">
          <div>Products: {spreadsheets.Products?.get('B9')} items, Avg: ${spreadsheets.Products?.get('B10')}</div>
          <div>Sales: ${spreadsheets.Sales?.get('B10')} revenue from {spreadsheets.Sales?.get('B11')} units</div>
          <div>Dashboard: Revenue per product: ${spreadsheets.Dashboard?.get('B14')}</div>
        </div>
      </div>
    </div>
  );
}
