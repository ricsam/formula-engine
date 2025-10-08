import {
  getCellReference,
  parseCellReference,
  Spreadsheet,
} from "@anocca-pub/components";
import React, { useEffect, useMemo, useState, useCallback } from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { FormulaEngine } from "../src/core/engine";
import type { SelectionManager } from "@ricsam/selection-manager";
import { createEngineWithMultiSheetData } from "./lib/multisheet-data";
import { useEngine } from "src/react/hooks";
import type { CellAddress } from "src/core/types";

interface SheetComponentProps {
  sheetName: "Dashboard" | "Sales" | "Products";
  workbookName: string;
  spreadsheetData: Map<string, any>;
  engine: FormulaEngine;
  activeSheet: string;
  selectedCell: string | null;
  onSheetActivate: (sheetName: "Dashboard" | "Sales" | "Products") => void;
  onCellSelect: (cell: string | null) => void;
}

function SheetComponent({
  sheetName,
  workbookName,
  spreadsheetData,
  engine,
  activeSheet,
  selectedCell,
  onSheetActivate,
  onCellSelect,
}: SheetComponentProps) {
  const selectionManagerEffects = useCallback(
    (selectionManager: SelectionManager) => {
      selectionManager.observeStateChange(
        (state) => {
          if (state.isSelecting?.type === "drag") {
            const cell = state.isSelecting.start;
            return getCellReference({
              rowIndex: cell.row,
              colIndex: cell.col,
            });
          }
          const cell = state.selections[state.selections.length - 1]?.start;
          if (cell) {
            return getCellReference({
              rowIndex: cell.row,
              colIndex: cell.col,
            });
          }
        },
        (cell) => {
          onCellSelect(cell ?? null);
          if (cell) {
            onSheetActivate(sheetName);
          }
        }
      );
    },
    [onCellSelect, onSheetActivate, sheetName]
  );

  return (
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
      <div className="border rounded-lg overflow-hidden bg-white flex-1">
        <Spreadsheet
          style={{ width: "100%", height: "100%" }}
          cellData={spreadsheetData}
          onCellDataChange={(updatedSpreadsheet) => {
            engine.setSheetContent(
              {
                sheetName,
                workbookName: "MultiSheetDemo",
              },
              updatedSpreadsheet
            );
          }}
          customCellRenderer={(cell) => {
            const value = engine.getCellValue({
              sheetName,
              colIndex: cell.colIndex,
              rowIndex: cell.rowIndex,
              workbookName,
            });
            return <div>{value}</div>;
          }}
          selection={{
            effects: selectionManagerEffects,
          }}
        />
      </div>
    </div>
  );
}

export function MultiSheetDemo() {
  const { engine, sheets } = useMemo(createEngineWithMultiSheetData, []);
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [activeSheet, setActiveSheet] = useState<
    "Dashboard" | "Sales" | "Products"
  >("Dashboard");
  const engineState = useEngine(engine);
  const spreadsheets = {
    Dashboard: engineState.workbooks.get("MultiSheetDemo")?.sheets.get(sheets.dashboard.name)!,
    Sales: engineState.workbooks.get("MultiSheetDemo")?.sheets.get(sheets.sales.name)!,
    Products: engineState.workbooks.get("MultiSheetDemo")?.sheets.get(sheets.products.name)!,
  };

  const cellSerialized = useMemo(() => {
    if (!selectedCell) {
      return;
    }
    const cellFormula = spreadsheets[activeSheet].content.get(selectedCell);
    return cellFormula;
  }, [activeSheet, selectedCell, spreadsheets]);

  const handleFormulaSubmit = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === "Enter" && selectedCell) {
        try {
          const { columnIndex, rowIndex } = parseCellReference(selectedCell);
          const address: CellAddress = {
            sheetName: activeSheet,
            colIndex: columnIndex,
            rowIndex: rowIndex,
            workbookName: "MultiSheetDemo",
          };

          engine.setCellContent(address, e.currentTarget.value);
        } catch (error) {
          console.error("Error updating cell:", error);
        }
      }
    },
    [selectedCell, activeSheet, engine]
  );

  const addNewSale = useCallback(() => {
    const salesData = engine.getSheetSerialized({ workbookName: "MultiSheetDemo", sheetName: sheets.sales.name });
    const salesKeys = Array.from(salesData.keys()).filter(
      (key) => key.startsWith("A") && key !== "A1"
    );
    const lastRow = Math.max(
      ...salesKeys.map((key) => parseInt(key.substring(1)))
    );
    const newRow = lastRow + 1;

    // Add a new sale with random product (now includes P006 and P007)
    const productIds = ["P001", "P002", "P003", "P004", "P005", "P006", "P007"];
    const randomProductId =
      productIds[Math.floor(Math.random() * productIds.length)];
    const randomQuantity = Math.floor(Math.random() * 5) + 1;

    const newSaleData = new Map<string, any>([
      [`A${newRow}`, `S${String(newRow - 1).padStart(3, "0")}`],
      [`B${newRow}`, "2024-01-25"],
      [`C${newRow}`, randomProductId],
      [`D${newRow}`, `=INDEX(Products!B:B,MATCH(C${newRow},Products!A:A,0))`],
      [`E${newRow}`, randomQuantity],
      [`F${newRow}`, `=INDEX(Products!D:D,MATCH(C${newRow},Products!A:A,0))`],
      [`G${newRow}`, `=E${newRow}*F${newRow}`],
      [`H${newRow}`, `=INDEX(Products!C:C,MATCH(C${newRow},Products!A:A,0))`],
    ]);

    newSaleData.forEach((value, key) => {
      const { columnIndex, rowIndex } = parseCellReference(key);
      engine.setCellContent(
        {
          sheetName: sheets.sales.name,
          colIndex: columnIndex,
          rowIndex: rowIndex,
          workbookName: "MultiSheetDemo",
        },
        value
      );
    });
  }, [engine, sheets.sales.name]);

  return (
    <div className="flex flex-col gap-4 h-full w-full p-8">
      {/* Header with controls */}
      <div className="flex items-center gap-4 p-4 border-b">
        <h2 className="text-xl font-bold">
          Multi-Sheet Demo - Cross-Sheet References
        </h2>
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
          defaultValue={cellSerialized ? String(cellSerialized) : ""}
          onKeyDown={handleFormulaSubmit}
          className="flex-1 font-mono"
          placeholder={
            selectedCell ? "Enter formula or value..." : "Select a cell to edit"
          }
          disabled={!selectedCell}
        />
      </div>

      {/* Three spreadsheets side by side */}
      <div className="flex-1 overflow-hidden px-4">
        <div className="grid grid-cols-3 gap-4 h-full">
          <SheetComponent
            sheetName="Products"
            spreadsheetData={spreadsheets.Products?.content ?? new Map()}
            engine={engine}
            activeSheet={activeSheet}
            selectedCell={selectedCell}
            onSheetActivate={setActiveSheet}
            onCellSelect={setSelectedCell}
            workbookName="MultiSheetDemo"
          />
          <SheetComponent
            sheetName="Sales"
            spreadsheetData={spreadsheets.Sales?.content ?? new Map()}
            engine={engine}
            activeSheet={activeSheet}
            selectedCell={selectedCell}
            onSheetActivate={setActiveSheet}
            onCellSelect={setSelectedCell}
            workbookName="MultiSheetDemo"
          />
          <SheetComponent
            sheetName="Dashboard"
            spreadsheetData={spreadsheets.Dashboard?.content ?? new Map()}
            engine={engine}
            activeSheet={activeSheet}
            selectedCell={selectedCell}
            onSheetActivate={setActiveSheet}
            onCellSelect={setSelectedCell}
            workbookName="MultiSheetDemo"
          />
        </div>
      </div>

      {/* Key features info */}
      <div className="px-4 py-2 bg-gray-50 border-t text-xs">
        <strong>Key Features Demonstrated:</strong>
        <span className="ml-2">
          Cross-sheet references (Products!A1, Sales!B10) • INDEX/MATCH lookups
          • SUMIF/COUNTIF aggregations • Text functions (CONCATENATE, UPPER) •
          Dynamic calculations • Real-time updates across sheets
        </span>
      </div>

      {/* Live summary */}
      <div className="px-4 py-2 bg-blue-50 text-xs">
        <strong>Live Summary:</strong>
        <div className="grid grid-cols-3 gap-4 mt-1">
          <div>
            Products: {spreadsheets.Dashboard?.content.get("B5")} items, Avg
            Price: ${spreadsheets.Dashboard?.content.get("B6")}
          </div>
          <div>
            Sales: ${spreadsheets.Dashboard?.content.get("B13")} revenue from{" "}
            {spreadsheets.Dashboard?.content.get("B14")} units
          </div>
          <div>
            Best Category: {spreadsheets.Dashboard?.content.get("E31")} leading
            with {spreadsheets.Dashboard?.content.get("B32")} market share
          </div>
        </div>
      </div>
    </div>
  );
}
