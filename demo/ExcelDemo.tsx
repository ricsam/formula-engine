import { useState, useCallback, useMemo, useEffect } from "react";
import { FormulaEngine } from "../src/core/engine";
import { useGlobalNamedExpressions, useSerializedSheet, useTables } from "../src/react/hooks";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Plus, X, Edit2, Check, X as Cancel, Save, Upload, Calculator, ChevronDown, ChevronUp } from "lucide-react";
import { SpreadsheetWithFormulaBar } from "./components/SpreadsheetWithFormulaBar";
import type { SpreadsheetRangeEnd, TableDefinition } from "src/core/types";

interface SheetTab {
  name: string;
}

type SerializedTableDefinition = {
  name: string;
  start: {
    rowIndex: number;
    colIndex: number;
  };
  headers: Record<string, { name: string; index: number }>;
  endRow: SpreadsheetRangeEnd;
  sheetName: string;
}

interface SavedSpreadsheetData {
  sheets: Array<{
    name: string;
    cells: Record<string, any>;
    namedExpressions: Record<string, { name: string; expression: string }>;
  }>;
  globalNamedExpressions: Record<string, { name: string; expression: string }>;
  tables: Record<string, SerializedTableDefinition>;
  activeSheet: string;
}

const STORAGE_KEY = "formula-engine-excel-demo";

const loadFromLocalStorage = (): SavedSpreadsheetData | null => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    return saved ? JSON.parse(saved) : null;
  } catch (error) {
    console.error("Failed to load from localStorage:", error);
    return null;
  }
};

const createEngine = () => {
  const engine = FormulaEngine.buildEmpty();
  const savedData = loadFromLocalStorage();

  if (savedData && savedData.sheets.length > 0) {
    // Load saved sheets
    const loadedSheets: string[] = [];
    for (const savedSheet of savedData.sheets) {
      const sheet = engine.addSheet(savedSheet.name);
      loadedSheets.push(sheet.name);

      // Build the cell content map for faster loading
      const cellContentMap = new Map<string, any>();
      for (const [cellId, value] of Object.entries(savedSheet.cells)) {
        if (value !== undefined) {
          cellContentMap.set(cellId, value);
        }
      }

      // Set all cell content at once (faster than individual setCellContent calls)
      engine.setSheetContent(sheet.name, cellContentMap);

      // Load named expressions for this sheet
      const namedExpressionsMap = new Map();
      for (const [name, expr] of Object.entries(savedSheet.namedExpressions || {})) {
        namedExpressionsMap.set(name, expr);
      }
      if (namedExpressionsMap.size > 0) {
        engine.setNamedExpressions(sheet.name, namedExpressionsMap);
      }

      // Tables are loaded globally after all sheets are processed
    }

    // Load global named expressions
    if (savedData.globalNamedExpressions) {
      const globalNamedExpressionsMap = new Map();
      for (const [name, expr] of Object.entries(savedData.globalNamedExpressions)) {
        globalNamedExpressionsMap.set(name, expr);
      }
      if (globalNamedExpressionsMap.size > 0) {
        engine.setGlobalNamedExpressions(globalNamedExpressionsMap);
      }
    }

    // Load global tables
    if (savedData.tables) {
      const tablesMap = new Map<string, TableDefinition>();
      for (const [name, serializedTable] of Object.entries(savedData.tables)) {
        // Convert serialized headers record back to Map
        const headersMap = new Map<string, { name: string; index: number }>();
        for (const [headerName, headerData] of Object.entries(serializedTable.headers)) {
          headersMap.set(headerName, headerData);
        }
        
        const table: TableDefinition = {
          ...serializedTable,
          headers: headersMap,
        };
        
        tablesMap.set(name, table);
      }
      if (tablesMap.size > 0) {
        engine.setTables(tablesMap);
      }
    }

    return {
      engine,
      sheetName: savedData.activeSheet || loadedSheets[0]!,
      savedSheets: loadedSheets,
    };
  } else {
    // Create first sheet with sample data
    const sheetName = engine.addSheet("Sheet1").name;
    return { engine, sheetName, savedSheets: [sheetName] };
  }
};

export function ExcelDemo() {
  const {
    engine,
    sheetName: initialSheetName,
    savedSheets,
  } = useMemo(() => createEngine(), []);

  const [sheets, setSheets] = useState<SheetTab[]>(
    savedSheets.map((name) => ({ name }))
  );

  const [activeSheet, setActiveSheet] = useState(initialSheetName);
  const [editingSheet, setEditingSheet] = useState<string | null>(null);
  const [editingName, setEditingName] = useState("");
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  
  // Named expressions UI state
  const [showNamedExpressions, setShowNamedExpressions] = useState(false);
  const [newExpressionName, setNewExpressionName] = useState("");
  const [newExpressionFormula, setNewExpressionFormula] = useState("");
  const [newExpressionScope, setNewExpressionScope] = useState<"global" | "sheet">("global");
  
  // Editing states for named expressions
  const [editingExpression, setEditingExpression] = useState<{name: string, isGlobal: boolean} | null>(null);
  const [editingExpressionName, setEditingExpressionName] = useState("");
  const [editingExpressionFormula, setEditingExpressionFormula] = useState("");
  
  // Table management UI state
  const [editingTable, setEditingTable] = useState<string | null>(null);
  const [editingTableName, setEditingTableName] = useState("");

  // Get named expressions and tables using hooks
  const globalNamedExpressions = useGlobalNamedExpressions(engine);
  const serializedSheet = useSerializedSheet(engine, activeSheet);
  const sheetNamedExpressions = serializedSheet.namedExpressions;
  const tables = useTables(engine);

  // Save to localStorage
  const saveToLocalStorage = useCallback(() => {
    try {
      const sheetsData = Array.from(engine.sheets.entries()).map(
        ([name, sheet]) => ({
          name,
          cells: Object.fromEntries(sheet.content),
          namedExpressions: Object.fromEntries(engine.getNamedExpressionsSerialized(name)),
        })
      );

      // Serialize tables with headers converted from Map to Record
      const serializedTables: Record<string, SerializedTableDefinition> = {};
      for (const [tableName, table] of engine.getTablesSerialized()) {
        serializedTables[tableName] = {
          ...table,
          headers: Object.fromEntries(table.headers),
        };
      }

      const dataToSave: SavedSpreadsheetData = {
        sheets: sheetsData,
        globalNamedExpressions: Object.fromEntries(engine.getGlobalNamedExpressionsSerialized()),
        tables: serializedTables,
        activeSheet,
      };

      localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToSave));
      setHasUnsavedChanges(false);
      console.log("Spreadsheet saved to localStorage");
    } catch (error) {
      console.error("Failed to save to localStorage:", error);
    }
  }, [engine, activeSheet]);

  // Mark as having unsaved changes when sheets change
  const markUnsavedChanges = useCallback(() => {
    setHasUnsavedChanges(true);
  }, []);

  // Auto-save when engine changes (optional)
  useEffect(() => {
    const unsubscribe = engine.onCellsUpdate(activeSheet, markUnsavedChanges);
    return unsubscribe;
  }, [engine, activeSheet, markUnsavedChanges]);

  // Track global named expression changes
  useEffect(() => {
    const unsubscribe = engine.on("global-named-expressions-updated", markUnsavedChanges);
    return unsubscribe;
  }, [engine, markUnsavedChanges]);

  // Track table changes
  useEffect(() => {
    const unsubscribe = engine.on("tables-updated", markUnsavedChanges);
    return unsubscribe;
  }, [engine, markUnsavedChanges]);

  // Add named expression
  const addNamedExpression = useCallback(() => {
    if (newExpressionName.trim() && newExpressionFormula.trim()) {
      engine.addNamedExpression({
        expressionName: newExpressionName.trim(),
        expression: newExpressionFormula.trim(),
        sheetName: newExpressionScope === "sheet" ? activeSheet : undefined,
      });
      
      setNewExpressionName("");
      setNewExpressionFormula("");
      markUnsavedChanges();
    }
  }, [newExpressionName, newExpressionFormula, newExpressionScope, activeSheet, engine, markUnsavedChanges]);

  // Delete named expression
  const deleteNamedExpression = useCallback((name: string, isGlobal: boolean) => {
    try {
      const success = engine.removeNamedExpression({
        expressionName: name,
        sheetName: isGlobal ? undefined : activeSheet,
      });
      
      if (success) {
        markUnsavedChanges();
      }
    } catch (error) {
      console.error("Failed to delete named expression:", error);
    }
  }, [engine, activeSheet, markUnsavedChanges]);

  // Start editing named expression
  const startEditingExpression = useCallback((name: string, expression: string, isGlobal: boolean) => {
    setEditingExpression({ name, isGlobal });
    setEditingExpressionName(name);
    setEditingExpressionFormula(expression);
  }, []);

  // Save named expression changes
  const saveExpressionChanges = useCallback(() => {
    if (!editingExpression || !editingExpressionName.trim() || !editingExpressionFormula.trim()) {
      return;
    }

    try {
      const { name: oldName, isGlobal } = editingExpression;
      const sheetName = isGlobal ? undefined : activeSheet;
      
      // If name changed, rename the expression
      if (oldName !== editingExpressionName.trim()) {
        engine.renameNamedExpression({
          expressionName: oldName,
          sheetName,
          newName: editingExpressionName.trim(),
        });
      }
      
      // Update the expression formula
      engine.updateNamedExpression({
        expressionName: editingExpressionName.trim(),
        expression: editingExpressionFormula.trim(),
        sheetName,
      });
      
      markUnsavedChanges();
      setEditingExpression(null);
      setEditingExpressionName("");
      setEditingExpressionFormula("");
    } catch (error) {
      console.error("Failed to update named expression:", error);
    }
  }, [editingExpression, editingExpressionName, editingExpressionFormula, activeSheet, engine, markUnsavedChanges]);

  // Cancel editing named expression
  const cancelEditingExpression = useCallback(() => {
    setEditingExpression(null);
    setEditingExpressionName("");
    setEditingExpressionFormula("");
  }, []);

  // Table management functions
  const deleteTable = useCallback((tableName: string) => {
    try {
      engine.removeTable({ tableName });
      markUnsavedChanges();
    } catch (error) {
      console.error("Failed to delete table:", error);
    }
  }, [engine, markUnsavedChanges]);

  // Start editing table name
  const startEditingTable = useCallback((tableName: string) => {
    setEditingTable(tableName);
    setEditingTableName(tableName);
  }, []);

  // Save table name changes
  const saveTableChanges = useCallback(() => {
    if (!editingTable || !editingTableName.trim()) {
      return;
    }

    try {
      if (editingTable !== editingTableName.trim()) {
        engine.renameTable({
          oldName: editingTable,
          newName: editingTableName.trim(),
        });
      }
      
      markUnsavedChanges();
      setEditingTable(null);
      setEditingTableName("");
    } catch (error) {
      console.error("Failed to rename table:", error);
    }
  }, [editingTable, editingTableName, engine, markUnsavedChanges]);

  // Cancel editing table
  const cancelEditingTable = useCallback(() => {
    setEditingTable(null);
    setEditingTableName("");
  }, []);

  // Add new sheet
  const addSheet = useCallback(() => {
    const newSheetCount = sheets.length + 1;
    const newSheetName = `Sheet${newSheetCount}`;
    const addedSheetName = engine.addSheet(newSheetName).name;

    const newSheet: SheetTab = {
      name: addedSheetName,
    };

    setSheets((prev) => [...prev, newSheet]);
    setActiveSheet(addedSheetName);
    markUnsavedChanges();
  }, [sheets.length, engine, markUnsavedChanges]);

  // Delete sheet
  const deleteSheet = useCallback(
    (sheetName: string) => {
      if (sheets.length <= 1) return; // Don't delete the last sheet

      try {
        engine.removeSheet(sheetName);
        
        setSheets((prev) => {
          const newSheets = prev.filter((sheet) => sheet.name !== sheetName);

          // If we deleted the active sheet, switch to the first remaining sheet
          if (sheetName === activeSheet && newSheets.length > 0) {
            setActiveSheet(newSheets[0]!.name);
          }

          return newSheets;
        });
        markUnsavedChanges();
      } catch (error) {
        console.error("Failed to delete sheet:", error);
      }
    },
    [sheets.length, activeSheet, engine, markUnsavedChanges]
  );

  // Start editing sheet name
  const startEditingSheet = useCallback(
    (sheetName: string, currentName: string) => {
      setEditingSheet(sheetName);
      setEditingName(currentName);
    },
    []
  );

  // Save sheet name
  const saveSheetName = useCallback(() => {
    if (editingSheet !== null && editingName.trim()) {
      try {
        engine.renameSheet(editingSheet, editingName.trim());

        setSheets((prev) =>
          prev.map((sheet) =>
            sheet.name === editingSheet
              ? { ...sheet, name: editingName.trim() }
              : sheet
          )
        );

        // Update active sheet name if we renamed the active sheet
        if (editingSheet === activeSheet) {
          setActiveSheet(editingName.trim());
        }

        markUnsavedChanges();
      } catch (error) {
        console.error("Failed to rename sheet:", error);
      }
    }

    setEditingSheet(null);
    setEditingName("");
  }, [editingSheet, editingName, activeSheet, engine, markUnsavedChanges]);

  // Cancel editing
  const cancelEditing = useCallback(() => {
    setEditingSheet(null);
    setEditingName("");
  }, []);

  // Handle key press in name input
  const handleKeyPress = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "Enter") {
        saveSheetName();
      } else if (e.key === "Escape") {
        cancelEditing();
      }
    },
    [saveSheetName, cancelEditing]
  );

  return (
    <div className="h-full flex flex-col">
      {/* Excel-style header */}
      <div className="border-b border-gray-200 bg-gray-50">
        <div className="p-3 border-b border-gray-200">
          <div className="flex items-center justify-between">
            <h1 className="text-lg font-semibold text-gray-800">
              FormulaEngine Excel Demo
            </h1>
            <div className="flex items-center gap-4 text-sm text-gray-600">
              <span data-testid="active-sheet-display">
                Active Sheet:{" "}
                <strong>
                  {sheets.find((s) => s.name === activeSheet)?.name}
                </strong>
              </span>
              <span data-testid="total-sheets-count">Total Sheets: {sheets.length}</span>
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  className="border-gray-300 text-gray-700"
                  onClick={() => setShowNamedExpressions(!showNamedExpressions)}
                  data-testid="named-expressions-toggle"
                >
                  <Calculator className="h-4 w-4 mr-1" />
                  Expressions & Tables
                  {showNamedExpressions ? (
                    <ChevronUp className="h-3 w-3 ml-1" />
                  ) : (
                    <ChevronDown className="h-3 w-3 ml-1" />
                  )}
                </Button>
                <Button
                  size="sm"
                  variant={hasUnsavedChanges ? "default" : "outline"}
                  className={`
                    ${
                      hasUnsavedChanges
                        ? "bg-blue-600 hover:bg-blue-700 text-white"
                        : "border-gray-300 text-gray-700"
                    }
                  `}
                  onClick={saveToLocalStorage}
                  data-testid="save-button"
                >
                  <Save className="h-4 w-4 mr-1" />
                  {hasUnsavedChanges ? "Save Changes" : "Saved"}
                </Button>
                {hasUnsavedChanges && (
                  <span className="text-xs text-orange-600 font-medium" data-testid="unsaved-changes-indicator">
                    Unsaved changes
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Named Expressions & Tables Panel */}
      {showNamedExpressions && (
        <div className="border-b border-gray-200 bg-gray-50 p-4">
          <div className="space-y-4">
            {/* Add New Named Expression */}
            <div className="bg-white p-3 rounded border border-gray-200">
              <h3 className="text-sm font-semibold text-gray-800 mb-3">Add Named Expression</h3>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">Name</label>
                  <Input
                    placeholder="e.g., TAX_RATE"
                    value={newExpressionName}
                    onChange={(e) => setNewExpressionName(e.target.value)}
                    className="text-sm"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">Formula</label>
                  <Input
                    placeholder="e.g., 0.08"
                    value={newExpressionFormula}
                    onChange={(e) => setNewExpressionFormula(e.target.value)}
                    className="text-sm"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">Scope</label>
                  <select
                    value={newExpressionScope}
                    onChange={(e) => setNewExpressionScope(e.target.value as "global" | "sheet")}
                    className="w-full p-2 text-sm border border-gray-300 rounded-md"
                  >
                    <option value="global">Global</option>
                    <option value="sheet">Current Sheet</option>
                  </select>
                </div>
                <div className="flex items-end">
                  <Button
                    size="sm"
                    onClick={addNamedExpression}
                    disabled={!newExpressionName.trim() || !newExpressionFormula.trim()}
                    className="w-full"
                  >
                    <Plus className="h-4 w-4 mr-1" />
                    Add
                  </Button>
                </div>
              </div>
            </div>

            {/* Existing Named Expressions and Tables */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Global Named Expressions */}
              <div className="bg-white p-3 rounded border border-gray-200">
                <h3 className="text-sm font-semibold text-gray-800 mb-3">
                  Global Named Expressions ({globalNamedExpressions.size})
                </h3>
                <div className="space-y-2 max-h-40 overflow-y-auto">
                  {globalNamedExpressions.size === 0 ? (
                    <p className="text-xs text-gray-500 italic">No global named expressions</p>
                  ) : (
                    Array.from(globalNamedExpressions.entries()).map(([name, expr]) => (
                      <div key={name} className="flex items-center justify-between bg-gray-50 p-2 rounded">
                        {editingExpression?.name === name && editingExpression?.isGlobal ? (
                          <div className="flex-1 flex items-center gap-2">
                            <div className="flex-1 space-y-1">
                              <Input
                                value={editingExpressionName}
                                onChange={(e) => setEditingExpressionName(e.target.value)}
                                className="h-6 text-xs"
                                placeholder="Name"
                              />
                              <Input
                                value={editingExpressionFormula}
                                onChange={(e) => setEditingExpressionFormula(e.target.value)}
                                className="h-6 text-xs"
                                placeholder="Formula"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-green-600 hover:text-green-700"
                                onClick={saveExpressionChanges}
                              >
                                <Check className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-gray-500 hover:text-gray-700"
                                onClick={cancelEditingExpression}
                              >
                                <Cancel className="h-3 w-3" />
                              </Button>
                            </div>
                          </div>
                        ) : (
                          <>
                            <div className="flex-1 min-w-0">
                              <div className="font-medium text-xs text-gray-800 truncate">{name}</div>
                              <div className="text-xs text-gray-600 truncate">{expr.expression}</div>
                            </div>
                            <div className="flex items-center gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-blue-500 hover:text-blue-700"
                                onClick={() => startEditingExpression(name, expr.expression, true)}
                                data-testid={`edit-global-named-expression-${name}`}
                              >
                                <Edit2 className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                                onClick={() => deleteNamedExpression(name, true)}
                                data-testid={`delete-global-named-expression-${name}`}
                              >
                                <X className="h-3 w-3" />
                              </Button>
                            </div>
                          </>
                        )}
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Sheet Named Expressions */}
              <div className="bg-white p-3 rounded border border-gray-200">
                <h3 className="text-sm font-semibold text-gray-800 mb-3">
                  Sheet Named Expressions ({sheetNamedExpressions.size})
                </h3>
                <div className="space-y-2 max-h-40 overflow-y-auto">
                  {sheetNamedExpressions.size === 0 ? (
                    <p className="text-xs text-gray-500 italic">No sheet named expressions</p>
                  ) : (
                    Array.from(sheetNamedExpressions.entries()).map(([name, expr]) => (
                      <div key={name} className="flex items-center justify-between bg-gray-50 p-2 rounded">
                        {editingExpression?.name === name && !editingExpression?.isGlobal ? (
                          <div className="flex-1 flex items-center gap-2">
                            <div className="flex-1 space-y-1">
                              <Input
                                value={editingExpressionName}
                                onChange={(e) => setEditingExpressionName(e.target.value)}
                                className="h-6 text-xs"
                                placeholder="Name"
                              />
                              <Input
                                value={editingExpressionFormula}
                                onChange={(e) => setEditingExpressionFormula(e.target.value)}
                                className="h-6 text-xs"
                                placeholder="Formula"
                              />
                            </div>
                            <div className="flex flex-col gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-green-600 hover:text-green-700"
                                onClick={saveExpressionChanges}
                              >
                                <Check className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-gray-500 hover:text-gray-700"
                                onClick={cancelEditingExpression}
                              >
                                <Cancel className="h-3 w-3" />
                              </Button>
                            </div>
                          </div>
                        ) : (
                          <>
                            <div className="flex-1 min-w-0">
                              <div className="font-medium text-xs text-gray-800 truncate">{name}</div>
                              <div className="text-xs text-gray-600 truncate">{expr.expression}</div>
                            </div>
                            <div className="flex items-center gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-blue-500 hover:text-blue-700"
                                onClick={() => startEditingExpression(name, expr.expression, false)}
                                data-testid={`edit-sheet-named-expression-${name}`}
                              >
                                <Edit2 className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                                onClick={() => deleteNamedExpression(name, false)}
                                data-testid={`delete-sheet-named-expression-${name}`}
                              >
                                <X className="h-3 w-3" />
                              </Button>
                            </div>
                          </>
                        )}
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Tables */}
              <div className="bg-white p-3 rounded border border-gray-200">
                <h3 className="text-sm font-semibold text-gray-800 mb-3">
                  Tables ({tables.size})
                </h3>
                <div className="space-y-2 max-h-40 overflow-y-auto">
                  {tables.size === 0 ? (
                    <p className="text-xs text-gray-500 italic">No tables</p>
                  ) : (
                    Array.from(tables.entries()).map(([name, table]) => (
                      <div key={name} className="flex items-center justify-between bg-gray-50 p-2 rounded">
                        {editingTable === name ? (
                          <div className="flex-1 flex items-center gap-2">
                            <Input
                              value={editingTableName}
                              onChange={(e) => setEditingTableName(e.target.value)}
                              className="h-6 text-xs flex-1"
                              placeholder="Table name"
                            />
                            <div className="flex items-center gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-green-600 hover:text-green-700"
                                onClick={saveTableChanges}
                              >
                                <Check className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-gray-500 hover:text-gray-700"
                                onClick={cancelEditingTable}
                              >
                                <Cancel className="h-3 w-3" />
                              </Button>
                            </div>
                          </div>
                        ) : (
                          <>
                            <div className="flex-1 min-w-0">
                              <div className="font-medium text-xs text-gray-800 truncate">{name}</div>
                              <div className="text-xs text-gray-600 truncate">
                                {table.sheetName} • {table.start.rowIndex + 1}:{String.fromCharCode(65 + table.start.colIndex)}
                              </div>
                            </div>
                            <div className="flex items-center gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-blue-500 hover:text-blue-700"
                                onClick={() => startEditingTable(name)}
                                data-testid={`edit-table-${name}`}
                              >
                                <Edit2 className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                                onClick={() => deleteTable(name)}
                                data-testid={`delete-table-${name}`}
                              >
                                <X className="h-3 w-3" />
                              </Button>
                            </div>
                          </>
                        )}
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Main spreadsheet area with formula bar */}
      <div className="flex-1 overflow-hidden">
        <SpreadsheetWithFormulaBar
          key={activeSheet} // Re-mount component when sheet changes
          sheetName={activeSheet}
          engine={engine}
          tables={tables}
          globalNamedExpressions={globalNamedExpressions}
        />
      </div>

      {/* Excel-style sheet tabs at bottom */}
      <div className="border-t border-gray-200 bg-gray-50 p-1 flex items-center gap-1">
        {/* Sheet tabs */}
        <div className="flex items-center gap-1 flex-1 overflow-x-auto">
          {sheets.map((sheet) => (
            <div
              key={sheet.name}
              className={`
                group relative flex items-center gap-1 px-3 py-1 border border-gray-300 rounded-t-md cursor-pointer
                ${
                  sheet.name === activeSheet
                    ? "bg-white border-b-white -mb-px z-10"
                    : "bg-gray-100 hover:bg-gray-200"
                }
              `}
              onClick={() => setActiveSheet(sheet.name)}
              data-testid={`sheet-tab-${sheet.name}`}
            >
              {editingSheet === sheet.name ? (
                <div className="flex items-center gap-1">
                  <Input
                    value={editingName}
                    onChange={(e) => setEditingName(e.target.value)}
                    onKeyDown={handleKeyPress}
                    onBlur={saveSheetName}
                    className="h-6 px-1 text-xs w-20 min-w-0"
                  />
                  <Button
                    size="sm"
                    variant="ghost"
                    className="h-4 w-4 p-0"
                    onClick={(e) => {
                      e.stopPropagation();
                      saveSheetName();
                    }}
                  >
                    <Check className="h-3 w-3" />
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    className="h-4 w-4 p-0"
                    onClick={(e) => {
                      e.stopPropagation();
                      cancelEditing();
                    }}
                  >
                    <Cancel className="h-3 w-3" />
                  </Button>
                </div>
              ) : (
                <>
                  <span className="text-xs font-medium text-gray-700 select-none">
                    {sheet.name}
                  </span>

                  {/* Edit button - only show on hover */}
                  <Button
                    size="sm"
                    variant="ghost"
                    className="h-4 w-4 p-0 opacity-0 group-hover:opacity-100 transition-opacity"
                    onClick={(e) => {
                      e.stopPropagation();
                      startEditingSheet(sheet.name, sheet.name);
                    }}
                    data-testid={`edit-sheet-${sheet.name}`}
                  >
                    <Edit2 className="h-3 w-3" />
                  </Button>

                  {/* Delete button - only show on hover and if not last sheet */}
                  {sheets.length > 1 && (
                    <Button
                      size="sm"
                      variant="ghost"
                      className="h-4 w-4 p-0 opacity-0 group-hover:opacity-100 transition-opacity text-red-500 hover:text-red-700"
                      onClick={(e) => {
                        e.stopPropagation();
                        deleteSheet(sheet.name);
                      }}
                      data-testid={`delete-sheet-${sheet.name}`}
                    >
                      <X className="h-3 w-3" />
                    </Button>
                  )}
                </>
              )}
            </div>
          ))}

          {/* Add sheet button */}
          <Button
            size="sm"
            variant="ghost"
            className="h-7 w-7 p-0 border border-gray-300 rounded-t-md bg-gray-100 hover:bg-gray-200"
            onClick={addSheet}
            data-testid="add-sheet-button"
          >
            <Plus className="h-4 w-4" />
          </Button>
        </div>

        {/* Sheet scroll buttons (placeholder for future enhancement) */}
        <div className="flex items-center gap-1 text-xs text-gray-500">
          {sheets.length} sheet{sheets.length !== 1 ? "s" : ""}
        </div>
      </div>
    </div>
  );
}
