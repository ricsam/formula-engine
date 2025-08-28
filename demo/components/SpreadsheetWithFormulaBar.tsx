import {
  Spreadsheet,
  getCellReference,
  parseCellReference,
} from "@anocca-pub/components";
import type { SelectionManager, SMArea } from "@ricsam/selection-manager";
import React, {
  useCallback,
  useMemo,
  useRef,
  useState,
  useEffect,
} from "react";
import { useSerializedSheet } from "src/react/hooks";
import { FormulaEngine } from "../../src/core/engine";
import { Input } from "../components/ui/input";
import { Button } from "../components/ui/button";
import { Card } from "../components/ui/card";
import type {
  CellAddress,
  NamedExpression,
  SerializedCellValue,
  TableDefinition,
} from "src/core/types";
import { indexToColumn } from "src/core/utils";

interface SpreadsheetWithFormulaBarProps {
  sheetName: string;
  engine: FormulaEngine;
  tables: Map<string, TableDefinition>;
  globalNamedExpressions: Map<string, NamedExpression>;
  verboseErrors?: boolean;
}

export function SpreadsheetWithFormulaBar({
  sheetName,
  engine,
  tables,
  globalNamedExpressions,
  verboseErrors = false,
}: SpreadsheetWithFormulaBarProps) {
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [selectedArea, setSelectedArea] = useState<SMArea | null>(null);
  const [newTableName, setNewTableName] = useState<string>("");
  const [tableCreationCounter, setTableCreationCounter] = useState(0);
  const formulaInputRef = useRef<HTMLInputElement>(null);

  const { sheet, namedExpressions } = useSerializedSheet(engine, sheetName);

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
  }, [selectedCell, sheetName, engine, sheet]);

  // Get existing table names to avoid duplicates
  const existingTableNames = useMemo(() => {
    return Array.from(tables.keys());
  }, [tables]);

  // Get table count directly from engine for more immediate updates
  const engineTableCount = useMemo(() => {
    return engine.getTablesSerialized().size;
  }, [engine, tables, tableCreationCounter]); // Include counter to force updates

  // Calculate the default table name based on current state
  const defaultTableName = useMemo(() => {
    if (currentSelectedTable) {
      return currentSelectedTable.name;
    } else {
      // Use the engine table count for more immediate updates
      const nextNumber = engineTableCount + 1;
      return `Table${nextNumber}`;
    }
  }, [currentSelectedTable, engineTableCount]);

  // Use a key to reset the input when the context changes
  const tableInputKey = useMemo(() => {
    return currentSelectedTable
      ? `table-${currentSelectedTable.name}`
      : `new-table-${defaultTableName}`;
  }, [currentSelectedTable, defaultTableName]);

  const tableColumnNames = useMemo(() => {
    if (!currentSelectedTable) return [];
    return Array.from(currentSelectedTable.headers.keys());
  }, [currentSelectedTable]);

  // Get tables for the current sheet
  const currentSheetTables = useMemo(() => {
    return Array.from(tables.entries()).filter(
      ([_, table]) => table.sheetName === sheetName
    );
  }, [tables, sheetName]);

  const addTableFromSelection = useCallback(() => {
    const effectiveTableName = (newTableName || defaultTableName).trim();
    if (!selectedArea || !effectiveTableName) {
      return;
    }
    if (selectedArea.end.col.type === "infinity") {
      return;
    }

    const trimmedName = effectiveTableName;

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
                value: selectedArea.end.row.value - selectedArea.start.row,
              },
        numCols: selectedArea.end.col.value - selectedArea.start.col + 1,
      });
      // Clear the input so it falls back to the calculated defaultTableName
      setNewTableName("");
      // Force recalculation of table count
      setTableCreationCounter((prev) => prev + 1);
    } catch (error) {
      console.error("Error creating table:", error);
    }
  }, [
    selectedArea,
    sheetName,
    engine,
    newTableName,
    defaultTableName,
    existingTableNames,
  ]);

  const cellSerialized = useMemo(() => {
    if (!selectedCell) {
      return;
    }
    const cellFormula = sheet.get(selectedCell);
    return cellFormula;
  }, [sheetName, selectedCell, sheet]);

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
            data-testid="formula-bar-input"
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
                      {currentSelectedTable.endRow.type === "number" ? (
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
                      ) : (
                        <>
                          :
                          {indexToColumn(
                            currentSelectedTable.start.colIndex +
                              currentSelectedTable.headers.size - 1
                          )}
                          ∞
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

          {/* Table Management / Create Table from Selection */}
          <div className="flex items-center gap-2">
            {currentSelectedTable ? (
              /* Table Management for Selected Table */
              <>
                <Input
                  key={tableInputKey}
                  value={newTableName || defaultTableName}
                  onChange={(e) => setNewTableName(e.target.value)}
                  placeholder="Table name"
                  className="w-32 h-8 text-sm"
                  data-testid="table-name-input"
                />
                <Button
                  onClick={() => {
                    if (
                      currentSelectedTable &&
                      newTableName.trim() &&
                      newTableName.trim() !== currentSelectedTable.name
                    ) {
                      try {
                        engine.renameTable({
                          oldName: currentSelectedTable.name,
                          newName: newTableName.trim(),
                        });
                      } catch (error) {
                        console.error("Error renaming table:", error);
                      }
                    }
                  }}
                  disabled={
                    !newTableName.trim() ||
                    newTableName.trim() === currentSelectedTable.name ||
                    existingTableNames.includes(newTableName.trim())
                  }
                  size="sm"
                  className="h-8"
                  variant="default"
                  data-testid="rename-table-button"
                >
                  Rename
                </Button>
                <Button
                  onClick={() => {
                    if (currentSelectedTable) {
                      try {
                        engine.removeTable({
                          tableName: currentSelectedTable.name,
                        });
                      } catch (error) {
                        console.error("Error removing table:", error);
                      }
                    }
                  }}
                  size="sm"
                  className="h-8"
                  variant="destructive"
                  data-testid="remove-table-button"
                >
                  Remove
                </Button>
              </>
            ) : selectedArea ? (
              /* Create Table from Selection */
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
                  key={tableInputKey}
                  value={newTableName || defaultTableName}
                  onChange={(e) => setNewTableName(e.target.value)}
                  placeholder="Table name"
                  className="w-32 h-8 text-sm"
                  data-testid="table-name-input"
                />
                <Button
                  onClick={addTableFromSelection}
                  disabled={
                    !selectedArea ||
                    !(newTableName || defaultTableName).trim() ||
                    existingTableNames.includes(
                      (newTableName || defaultTableName).trim()
                    ) ||
                    selectedArea.end.col.type === "infinity"
                  }
                  size="sm"
                  className="h-8"
                  variant={
                    (newTableName || defaultTableName).trim() &&
                    existingTableNames.includes(
                      (newTableName || defaultTableName).trim()
                    )
                      ? "destructive"
                      : "default"
                  }
                  data-testid="create-table-button"
                >
                  {(newTableName || defaultTableName).trim() &&
                  existingTableNames.includes(
                    (newTableName || defaultTableName).trim()
                  )
                    ? "Name Exists"
                    : "Create Table"}
                </Button>
              </>
            ) : null}
          </div>
        </div>
      </div>

      {/* Main spreadsheet area */}
      <div className="flex-1 overflow-hidden">
        <Spreadsheet
          style={{ height: "100%", width: "100%" }}
          cellData={sheet as Map<string, string | number>}
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
            }, verboseErrors);

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
