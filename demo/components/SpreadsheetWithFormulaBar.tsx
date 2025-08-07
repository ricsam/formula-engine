import React, { useState, useCallback, useMemo, useRef } from 'react';
import { Spreadsheet, getCellReference, parseCellReference } from '@anocca-pub/components';
import { FormulaEngine } from '../../src/core/engine';
import { Input } from '../components/ui/input';
import type { SelectionManager } from "@ricsam/selection-manager";

interface SpreadsheetWithFormulaBarProps {
  sheetId: number;
  engine: FormulaEngine;
  onSheetDataChange: (sheetId: number, newData: Map<string, any>) => void;
}

export function SpreadsheetWithFormulaBar({ 
  sheetId, 
  engine, 
  onSheetDataChange 
}: SpreadsheetWithFormulaBarProps) {
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);

  // Get current sheet data for display
  const currentSheetData = useMemo(() => {
    const serializedData = engine.getSheetSerialized(sheetId);
    const filteredData = new Map<string, string | number>();
    for (const [key, value] of serializedData.entries()) {
      if (value !== undefined && value !== null) {
        if (typeof value === 'string' || typeof value === 'number') {
          filteredData.set(key, value);
        } else {
          filteredData.set(key, String(value));
        }
      }
    }
    return filteredData;
  }, [sheetId, engine]);

  // Get the serialized value of the selected cell for the formula bar
  const cellSerialized = useMemo(() => {
    if (!selectedCell) {
      return;
    }
    const { columnIndex, rowIndex } = parseCellReference(selectedCell);
    const cellFormula = engine.getCellSerialized({
      sheet: sheetId,
      col: columnIndex,
      row: rowIndex,
    });
    return cellFormula;
  }, [sheetId, selectedCell, engine]);

  // Handle cell data changes from the spreadsheet
  const onCellDataChange = useCallback((updatedSpreadsheet: Map<string, string | number>) => {
    engine.setSheetContents(sheetId, updatedSpreadsheet);
    onSheetDataChange(sheetId, engine.getSheetSerialized(sheetId));
  }, [sheetId, engine, onSheetDataChange]);

  // Handle formula submission from formula bar
  const handleFormulaSubmit = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === "Enter" && selectedCell) {
        try {
          const { columnIndex, rowIndex } = parseCellReference(selectedCell);
          const address = {
            sheet: sheetId,
            col: columnIndex,
            row: rowIndex,
          };

          engine.setCellContents(address, e.currentTarget.value);
          
          // Update the sheet data
          onSheetDataChange(sheetId, engine.getSheetSerialized(sheetId));
        } catch (error) {
          console.error("Error updating cell:", error);
        }
      }
    },
    [selectedCell, sheetId, engine, onSheetDataChange]
  );

  // Selection manager effects for tracking cell selection
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
          setSelectedCell(cell ?? null);
        }
      );
    },
    [setSelectedCell]
  );

  return (
    <div className="flex flex-col h-full">
      {/* Formula Bar Area */}
      <div className="p-2 bg-white border-b border-gray-200">
        <div className="flex items-center gap-2 text-sm text-gray-600">
          <span className="font-medium">Name Box:</span>
          <div className="px-2 py-1 border border-gray-300 rounded bg-gray-50 min-w-[80px] text-center">
            {selectedCell || 'A1'}
          </div>
          <span className="font-medium ml-4">Formula Bar:</span>
          <Input
            ref={formulaInputRef}
            key={selectedCell} // Force re-render when cell changes
            defaultValue={cellSerialized ? String(cellSerialized) : ""}
            onKeyDown={handleFormulaSubmit}
            className="flex-1 font-mono"
            placeholder={
              selectedCell ? "Enter formula or value..." : "Select a cell to edit"
            }
            disabled={!selectedCell}
          />
        </div>
      </div>
      
      {/* Main spreadsheet area */}
      <div className="flex-1 overflow-hidden">
        <Spreadsheet
          style={{ height: "100%", width: "100%" }}
          cellData={currentSheetData}
          onCellDataChange={onCellDataChange}
          selection={{
            effects: selectionManagerEffects,
          }}
          customCellRenderer={(cell) => {
            const value = engine.getCellValue({
              sheet: sheetId,
              col: cell.colIndex,
              row: cell.rowIndex,
            });
            
            if (typeof value === 'number') {
              // Format numbers nicely
              return (
                <div>
                  {value.toLocaleString(undefined, { 
                    minimumFractionDigits: 0, 
                    maximumFractionDigits: 2 
                  })}
                </div>
              );
            }
            
            return <div>{value?.toString() || ''}</div>;
          }}
        />
      </div>
    </div>
  );
}
