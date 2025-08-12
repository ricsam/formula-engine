import {
  Spreadsheet,
  getCellReference,
  parseCellReference,
} from "@anocca-pub/components";
import type { SelectionManager, SMArea } from "@ricsam/selection-manager";
import React, { useCallback, useMemo, useRef, useState } from "react";
import { useSerializedSheet } from "src/react/hooks";
import { FormulaEngine } from "../../src/core/engine";
import { Input } from "../components/ui/input";
import { Button } from "../components/ui/button";
import { Card } from "../components/ui/card";
import type { CellAddress, SerializedCellValue } from "src/core/types";

interface SpreadsheetWithFormulaBarProps {
  sheetName: string;
  engine: FormulaEngine;
}

export function SpreadsheetWithFormulaBar({
  sheetName,
  engine,
}: SpreadsheetWithFormulaBarProps) {
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [selectedArea, setSelectedArea] = useState<SMArea | null>(null);
  const [newTableName, setNewTableName] = useState<string>("Table1");
  const formulaInputRef = useRef<HTMLInputElement>(null);

  const currentSheetData = useSerializedSheet(engine, sheetName) as Map<
    string,
    string | number
  >;

  const currentSelectedTable = useMemo(() => {
    if (!selectedCell) {
      return undefined;
    }
    const parsed = parseCellReference(selectedCell);
    return engine.isCellInTable({
      sheetName,
      colIndex: parsed.columnIndex,
      rowIndex: parsed.rowIndex,
    });
  }, [selectedCell, sheetName, engine, currentSheetData]);

  console.log("@currentSelectedTable", currentSelectedTable, engine.tables);

  const tableColumnNames = useMemo(() => {
    if (!currentSelectedTable) return [];
    return Array.from(currentSelectedTable.headers.keys());
  }, [currentSelectedTable]);

  // Get existing table names to avoid duplicates
  const existingTableNames = useMemo(() => {
    const tables = engine.tables.get(sheetName);
    return tables ? Array.from(tables.keys()) : [];
  }, [engine, sheetName, currentSheetData]); // Include currentSheetData to refresh when tables change

  const addTableFromSelection = useCallback(() => {
    if (!selectedArea || !newTableName.trim()) {
      return;
    }
    if (selectedArea.end.col.type === "infinity") {
      return;
    }

    const trimmedName = newTableName.trim();

    // Check for duplicate names
    if (existingTableNames.includes(trimmedName)) {
      console.error(`Table name "${trimmedName}" already exists`);
      return;
    }

    try {
      engine.addTable({
        tableName: trimmedName,
        sheetName,
        start: getCellReference({
          rowIndex: selectedArea.start.row,
          colIndex: selectedArea.start.col,
        }),
        numRows:
          selectedArea.end.row.type === "infinity"
            ? { type: "infinity", sign: "positive" }
            : {
                type: "number",
                value: selectedArea.end.row.value - selectedArea.start.row + 1,
              },
        numCols: selectedArea.end.col.value - selectedArea.start.col + 1,
      });
      // Auto-increment table name for next table
      const match = trimmedName.match(/^(.+?)(\d+)$/);
      if (match && match[1] && match[2]) {
        const prefix = match[1];
        const num = match[2];
        setNewTableName(`${prefix}${parseInt(num) + 1}`);
      } else {
        setNewTableName(`${trimmedName}2`);
      }
    } catch (error) {
      console.error("Error creating table:", error);
    }
  }, [selectedArea, sheetName, engine, newTableName, existingTableNames]);

  const cellSerialized = useMemo(() => {
    if (!selectedCell) {
      return;
    }
    const cellFormula = currentSheetData.get(selectedCell);
    return cellFormula;
  }, [sheetName, selectedCell, currentSheetData]);

  // Handle cell data changes from the spreadsheet
  const onCellDataChange = useCallback(
    (updatedSpreadsheet: Map<string, string | number>) => {
      console.log("onCellDataChange", sheetName, updatedSpreadsheet);
      const data = new Map<string, SerializedCellValue>(updatedSpreadsheet);
      data.forEach((value, key) => {
        if (typeof value === "string") {
          const numberResult = isNumber(value);
          if (numberResult.isNumber) {
            data.set(key, numberResult.value);
          } else {
            const booleanResult = isBoolean(value);
            if (booleanResult.isBoolean) {
              data.set(key, booleanResult.value);
            }
          }
        }
      });
      engine.setSheetContent(sheetName, data);
    },
    [sheetName, engine]
  );

  // Handle formula submission from formula bar
  const handleFormulaSubmit = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === "Enter" && selectedCell) {
        try {
          const { columnIndex, rowIndex } = parseCellReference(selectedCell);
          const address: CellAddress = {
            sheetName,
            colIndex: columnIndex,
            rowIndex: rowIndex,
          };

          engine.setCellContent(address, e.currentTarget.value);
        } catch (error) {
          console.error("Error updating cell:", error);
        }
      }
    },
    [selectedCell, sheetName, engine]
  );

  // Selection manager effects for tracking cell selection
  const selectionManagerEffects = useCallback(
    (selectionManager: SelectionManager) => {
      const cleanups = [
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
        ),
        selectionManager.observeStateChange(
          (state) => {
            const currentSelection: SMArea | undefined =
              state.selections.length === 1 ? state.selections[0] : undefined;
            return currentSelection;
          },
          (selection) => {
            setSelectedArea(selection ?? null);
          }
        ),
      ];
      return () => {
        cleanups.forEach((cleanup) => cleanup());
      };
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
            {selectedCell || "A1"}
          </div>
          <span className="font-medium ml-4">Formula Bar:</span>
          <Input
            ref={formulaInputRef}
            key={selectedCell} // Force re-render when cell changes
            defaultValue={cellSerialized ? String(cellSerialized) : ""}
            onKeyDown={handleFormulaSubmit}
            className="flex-1 font-mono"
            placeholder={
              selectedCell
                ? "Enter formula or value..."
                : "Select a cell to edit"
            }
            disabled={!selectedCell}
          />
        </div>
      </div>

      {/* Table Management Panel */}
      <div className="py-3 bg-gray-50 border-b border-gray-200 h-18 flex items-center space-between w-full">
        <div className="flex items-center justify-between gap-4 w-full">
          {/* Current Table Info */}
          <div className="flex items-center gap-3">
            {currentSelectedTable ? (
              <Card className="px-3 py-2 bg-blue-50 border-blue-200">
                <div className="flex flex-col gap-1">
                  <div className="flex items-center gap-2 text-sm">
                    <span className="font-medium text-blue-800">Table:</span>
                    <span className="text-blue-700 font-mono">
                      {currentSelectedTable.name}
                    </span>
                    <span className="text-gray-500">•</span>
                    <span className="text-gray-600">
                      {getCellReference({
                        rowIndex: currentSelectedTable.start.rowIndex,
                        colIndex: currentSelectedTable.start.colIndex,
                      })}
                      {currentSelectedTable.endRow.type === "number" && (
                        <>
                          :
                          {getCellReference({
                            rowIndex:
                              currentSelectedTable.start.rowIndex +
                              currentSelectedTable.endRow.value -
                              1,
                            colIndex:
                              currentSelectedTable.start.colIndex +
                              currentSelectedTable.headers.size -
                              1,
                          })}
                        </>
                      )}
                    </span>
                    <span className="text-gray-500">•</span>
                    <span className="text-gray-600">
                      {currentSelectedTable.headers.size} columns
                    </span>
                    {currentSelectedTable.endRow.type === "number" && (
                      <>
                        <span className="text-gray-500">•</span>
                        <span className="text-gray-600">
                          {currentSelectedTable.endRow.value} rows
                        </span>
                      </>
                    )}
                  </div>
                  {tableColumnNames.length > 0 && (
                    <div className="flex items-center gap-1 text-xs text-gray-500">
                      <span>Columns:</span>
                      <div className="flex gap-1 flex-wrap">
                        {tableColumnNames.slice(0, 5).map((colName, idx) => (
                          <span
                            key={idx}
                            className="px-1 py-0.5 bg-gray-200 rounded text-gray-700"
                          >
                            {colName}
                          </span>
                        ))}
                        {tableColumnNames.length > 5 && (
                          <span className="text-gray-400">
                            +{tableColumnNames.length - 5} more
                          </span>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              </Card>
            ) : (
              <span className="text-sm text-gray-500">No table selected</span>
            )}
          </div>

          {/* Create Table from Selection */}
          <div className="flex items-center gap-2">
            {selectedArea && (
              <>
                <span className="text-sm text-gray-600">
                  Selection:{" "}
                  {getCellReference({
                    rowIndex: selectedArea.start.row,
                    colIndex: selectedArea.start.col,
                  })}
                  :
                  {selectedArea.end.row.type === "infinity" ||
                  selectedArea.end.col.type === "infinity"
                    ? "∞" // Show infinity symbol for infinite ranges
                    : getCellReference({
                        rowIndex:
                          selectedArea.end.row.type === "number"
                            ? selectedArea.end.row.value
                            : 0,
                        colIndex:
                          selectedArea.end.col.type === "number"
                            ? selectedArea.end.col.value
                            : 0,
                      })}
                  (
                  {selectedArea.end.col.type === "infinity"
                    ? "∞"
                    : selectedArea.end.col.type === "number"
                      ? selectedArea.end.col.value - selectedArea.start.col + 1
                      : 0}{" "}
                  ×{" "}
                  {selectedArea.end.row.type === "infinity"
                    ? "∞"
                    : selectedArea.end.row.type === "number"
                      ? selectedArea.end.row.value - selectedArea.start.row + 1
                      : 0}
                  )
                </span>
                <Input
                  value={newTableName}
                  onChange={(e) => setNewTableName(e.target.value)}
                  placeholder="Table name"
                  className="w-32 h-8 text-sm"
                />
                <Button
                  onClick={addTableFromSelection}
                  disabled={
                    !selectedArea ||
                    !newTableName.trim() ||
                    existingTableNames.includes(newTableName.trim()) ||
                    selectedArea.end.col.type === "infinity"
                  }
                  size="sm"
                  className="h-8"
                  variant={
                    newTableName.trim() &&
                    existingTableNames.includes(newTableName.trim())
                      ? "destructive"
                      : "default"
                  }
                >
                  {newTableName.trim() &&
                  existingTableNames.includes(newTableName.trim())
                    ? "Name Exists"
                    : "Create Table"}
                </Button>
              </>
            )}
          </div>
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
          customCellStyle={(cell) => {
            const tableInfo = engine.isCellInTable({
              sheetName,
              colIndex: cell.colIndex,
              rowIndex: cell.rowIndex,
            });

            if (!tableInfo) {
              return {}; // Not in a table, no special styling
            }

            const isHeaderRow = cell.rowIndex === tableInfo.start.rowIndex;
            const isFirstColumn = cell.colIndex === tableInfo.start.colIndex;
            const isLastColumn =
              cell.colIndex ===
              tableInfo.start.colIndex + tableInfo.headers.size - 1;

            // Calculate if this is the last row of the table
            const isLastRow =
              tableInfo.endRow.type === "number"
                ? cell.rowIndex ===
                  tableInfo.start.rowIndex + tableInfo.endRow.value - 1
                : false; // For infinite tables, we don't style the last row differently

            // Excel-like table styling
            const style: React.CSSProperties = {
              border: "1px solid #d0d7de",
            };

            if (isHeaderRow) {
              // Header row styling - blue theme like Excel
              style.backgroundColor = "#4472c4";
              style.color = "white";
              style.fontWeight = "bold";
              style.borderBottom = "2px solid #2f5597";
            } else {
              // Data rows - alternating background
              const dataRowIndex = cell.rowIndex - tableInfo.start.rowIndex - 1;
              if (dataRowIndex % 2 === 0) {
                style.backgroundColor = "#f8f9fa"; // Light gray for even rows
              } else {
                style.backgroundColor = "white"; // White for odd rows
              }
            }

            // Border styling
            if (isFirstColumn) {
              style.borderLeft = "2px solid #4472c4";
            }
            if (isLastColumn) {
              style.borderRight = "2px solid #4472c4";
            }
            if (isHeaderRow) {
              style.borderTop = "2px solid #4472c4";
            }
            if (isLastRow) {
              style.borderBottom = "2px solid #4472c4";
            }

            return style;
          }}
          customCellRenderer={(cell) => {
            const value = engine.getCellValue({
              sheetName,
              colIndex: cell.colIndex,
              rowIndex: cell.rowIndex,
            });

            if (typeof value === "number") {
              // Format numbers nicely
              return (
                <div>
                  {value.toLocaleString(undefined, {
                    minimumFractionDigits: 0,
                    maximumFractionDigits: 2,
                  })}
                </div>
              );
            }

            return <div>{value?.toString() || ""}</div>;
          }}
        />
      </div>
    </div>
  );
}

function isNumber(
  value: string
): { isNumber: true; value: number } | { isNumber: false } {
  // Empty string should not be treated as a number
  if (value === "") {
    return { isNumber: false };
  }

  // Handle comma as decimal separator (European style)
  const normalizedValue = value.replace(",", ".");

  // Check if it's a valid number
  const parsed = parseFloat(normalizedValue);

  // Make sure the entire string was consumed (no trailing characters)
  // and that it's a valid number
  if (
    !isNaN(parsed) &&
    isFinite(parsed) &&
    normalizedValue === String(parsed)
  ) {
    return { isNumber: true, value: parsed };
  }

  return { isNumber: false };
}

function isBoolean(
  value: string
): { isBoolean: true; value: boolean } | { isBoolean: false } {
  const lowerValue = value.toLowerCase().trim();

  // True values
  if (lowerValue === "true" || lowerValue === "yes" || lowerValue === "1") {
    return { isBoolean: true, value: true };
  }

  // False values
  if (lowerValue === "false" || lowerValue === "no" || lowerValue === "0") {
    return { isBoolean: true, value: false };
  }

  return { isBoolean: false };
}
