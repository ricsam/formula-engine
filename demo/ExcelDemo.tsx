import { useState, useCallback, useMemo, useEffect } from "react";
import { FormulaEngine } from "../src/core/engine";
import {
  useGlobalNamedExpressions,
  useSerializedSheet,
  useTables,
} from "../src/react/hooks";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import {
  Plus,
  X,
  Edit2,
  Check,
  X as Cancel,
  Save,
  Upload,
  Calculator,
  ChevronDown,
  ChevronUp,
  Bug,
  AlertTriangle,
} from "lucide-react";
import { SpreadsheetWithFormulaBar } from "./components/SpreadsheetWithFormulaBar";
import { Spreadsheet, Grid } from "@anocca-pub/components";
import type { GridChild } from "@anocca-pub/components";
import type { SpreadsheetRangeEnd, TableDefinition } from "src/core/types";


type SerializedTableDefinition = {
  name: string;
  start: {
    rowIndex: number;
    colIndex: number;
  };
  headers: [string, { name: string; index: number }][];
  endRow: SpreadsheetRangeEnd;
  sheetName: string;
  workbookName: string;
};

interface WorkbookGridItem {
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  activeSheet: string;
  sheets: Array<{
    name: string;
    cells: [string, any][];
    namedExpressions: [string, { name: string; expression: string }][];
  }>;
}

interface SavedSpreadsheetData {
  workbooks: WorkbookGridItem[];
  globalNamedExpressions: [string, { name: string; expression: string }][];
  tables: [string, SerializedTableDefinition][];
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

  if (savedData && savedData.workbooks.length > 0) {
    // Load saved workbooks and sheets
    for (const savedWorkbook of savedData.workbooks) {
      // Add workbook
      engine.addWorkbook(savedWorkbook.name);

      // Load sheets for this workbook
      for (const savedSheet of savedWorkbook.sheets) {
        const sheet = engine.addSheet({
          workbookName: savedWorkbook.name,
          sheetName: savedSheet.name,
        });

        // Build the cell content map for faster loading
        const cellContentMap = new Map<string, any>();
        for (const [cellId, value] of savedSheet.cells) {
          if (value !== undefined) {
            cellContentMap.set(cellId, value);
          }
        }

        // Set all cell content at once (faster than individual setCellContent calls)
        engine.setSheetContent(
          { workbookName: savedWorkbook.name, sheetName: sheet.name },
          cellContentMap
        );

        // Load named expressions for this sheet
        const namedExpressionsMap = new Map();
        for (const [name, expr] of savedSheet.namedExpressions || []) {
          namedExpressionsMap.set(name, expr);
        }
        if (namedExpressionsMap.size > 0) {
          engine.setNamedExpressions({
            workbookName: savedWorkbook.name,
            sheetName: sheet.name,
            expressions: namedExpressionsMap,
          });
        }
      }
    }

    // Load global named expressions
    if (savedData.globalNamedExpressions) {
      const globalNamedExpressionsMap = new Map();
      for (const [name, expr] of savedData.globalNamedExpressions) {
        globalNamedExpressionsMap.set(name, expr);
      }
      if (globalNamedExpressionsMap.size > 0) {
        engine.setGlobalNamedExpressions(globalNamedExpressionsMap);
      }
    }

    // Load global tables
    if (savedData.tables) {
      const tablesMap = new Map<string, TableDefinition>();
      for (const [name, serializedTable] of savedData.tables) {
        // Convert serialized headers array back to Map
        const headersMap = new Map<string, { name: string; index: number }>();
        for (const [headerName, headerData] of serializedTable.headers) {
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
      workbookGridItems: savedData.workbooks,
    };
  } else {
    // Create first workbook and sheet with sample data
    const workbookName = "Workbook1";
    engine.addWorkbook(workbookName);
    const sheetName = engine.addSheet({
      workbookName,
      sheetName: "Sheet1",
    }).name;
    
    const defaultWorkbookItem: WorkbookGridItem = {
      name: workbookName,
      x: 100,
      y: 100,
      width: 800,
      height: 600,
      activeSheet: sheetName,
      sheets: [{
        name: sheetName,
        cells: [],
        namedExpressions: [],
      }],
    };
    
    return { 
      engine, 
      workbookGridItems: [defaultWorkbookItem],
    };
  }
};

export function ExcelDemo() {
  const {
    engine,
    workbookGridItems: initialWorkbookGridItems,
  } = useMemo(() => createEngine(), []);

  const [workbookGridItems, setWorkbookGridItems] = useState<WorkbookGridItem[]>(initialWorkbookGridItems);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [debugMode, setDebugMode] = useState(false);
  const [verboseErrors, setVerboseErrors] = useState(false);

  // Named expressions UI state
  const [showNamedExpressions, setShowNamedExpressions] = useState(false);
  const [newExpressionName, setNewExpressionName] = useState("");
  const [newExpressionFormula, setNewExpressionFormula] = useState("");
  const [newExpressionScope, setNewExpressionScope] = useState<
    "global" | "sheet"
  >("global");
  const [selectedWorkbookForExpression, setSelectedWorkbookForExpression] = useState<string>("");

  // Editing states for named expressions
  const [editingExpression, setEditingExpression] = useState<{
    name: string;
    isGlobal: boolean;
    workbookName?: string;
  } | null>(null);
  const [editingExpressionName, setEditingExpressionName] = useState("");
  const [editingExpressionFormula, setEditingExpressionFormula] = useState("");

  // Table management UI state
  const [editingTable, setEditingTable] = useState<string | null>(null);
  const [editingTableName, setEditingTableName] = useState("");

  // Get named expressions and tables using hooks
  const globalNamedExpressions = useGlobalNamedExpressions(engine);
  const tables = useTables(engine);

  // Save to localStorage
  const saveToLocalStorage = useCallback(() => {
    try {
      // Serialize workbooks with their grid positions and sheets
      const workbooksData: WorkbookGridItem[] = workbookGridItems.map((gridItem) => {
        const workbook = engine.getWorkbooks().get(gridItem.name);
        if (!workbook) {
          throw new Error(`Workbook ${gridItem.name} not found in engine`);
        }

        return {
          ...gridItem,
          sheets: Array.from(workbook.sheets.entries()).map(([sheetName, sheet]) => ({
            name: sheetName,
            cells: Array.from(sheet.content.entries()),
            namedExpressions: Array.from(
              engine.getSheetExpressionsSerialized({ sheetName, workbookName: gridItem.name }).entries()
            ),
          })),
        };
      });

      // Serialize tables with headers converted from Map to array
      const serializedTables: [string, SerializedTableDefinition][] = [];
      for (const [tableName, table] of engine.getTablesSerialized()) {
        serializedTables.push([
          tableName,
          {
            ...table,
            headers: Array.from(table.headers.entries()),
          },
        ]);
      }

      const dataToSave: SavedSpreadsheetData = {
        workbooks: workbooksData,
        globalNamedExpressions: Array.from(
          engine.getGlobalNamedExpressionsSerialized().entries()
        ),
        tables: serializedTables,
      };

      localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToSave));
      setHasUnsavedChanges(false);
      console.log("Spreadsheet saved to localStorage");
    } catch (error) {
      console.error("Failed to save to localStorage:", error);
    }
  }, [engine, workbookGridItems]);

  // Mark as having unsaved changes when sheets change
  const markUnsavedChanges = useCallback(() => {
    setHasUnsavedChanges(true);
  }, []);

  // Track global named expression changes
  useEffect(() => {
    const unsubscribe = engine.on(
      "global-named-expressions-updated",
      markUnsavedChanges
    );
    return unsubscribe;
  }, [engine, markUnsavedChanges]);

  // Track table changes
  useEffect(() => {
    const unsubscribe = engine.on("tables-updated", markUnsavedChanges);
    return unsubscribe;
  }, [engine, markUnsavedChanges]);

  // Add new workbook
  const addWorkbook = useCallback(() => {
    const newWorkbookCount = workbookGridItems.length + 1;
    const newWorkbookName = `Workbook${newWorkbookCount}`;
    const newSheetName = "Sheet1";
    
    // Add to engine
    engine.addWorkbook(newWorkbookName);
    engine.addSheet({
      workbookName: newWorkbookName,
      sheetName: newSheetName,
    });

    // Add to grid
    const newGridItem: WorkbookGridItem = {
      name: newWorkbookName,
      x: 100 + (newWorkbookCount - 1) * 50, // Offset new workbooks
      y: 100 + (newWorkbookCount - 1) * 50,
      width: 800,
      height: 600,
      activeSheet: newSheetName,
      sheets: [{
        name: newSheetName,
        cells: [],
        namedExpressions: [],
      }],
    };

    setWorkbookGridItems(prev => [...prev, newGridItem]);
    markUnsavedChanges();
  }, [workbookGridItems.length, engine, markUnsavedChanges]);

  // Update workbook grid position
  const updateWorkbookPosition = useCallback((workbookName: string, x: number, y: number, width: number, height: number) => {
    setWorkbookGridItems(prev => 
      prev.map(item => 
        item.name === workbookName 
          ? { ...item, x, y, width, height }
          : item
      )
    );
    markUnsavedChanges();
  }, [markUnsavedChanges]);

  // Update active sheet for a workbook
  const updateWorkbookActiveSheet = useCallback((workbookName: string, sheetName: string) => {
    setWorkbookGridItems(prev => 
      prev.map(item => 
        item.name === workbookName 
          ? { ...item, activeSheet: sheetName }
          : item
      )
    );
    markUnsavedChanges();
  }, [markUnsavedChanges]);

  // Add named expression
  const addNamedExpression = useCallback(() => {
    if (newExpressionName.trim() && newExpressionFormula.trim()) {
      const workbookName = selectedWorkbookForExpression || workbookGridItems[0]?.name;
      const activeSheet = workbookGridItems.find(w => w.name === workbookName)?.activeSheet;
      
      engine.addNamedExpression({
        expressionName: newExpressionName.trim(),
        expression: newExpressionFormula.trim(),
        sheetName: newExpressionScope === "sheet" ? activeSheet : undefined,
        workbookName: newExpressionScope === "sheet" ? workbookName : undefined,
      });

      setNewExpressionName("");
      setNewExpressionFormula("");
      markUnsavedChanges();
    }
  }, [
    newExpressionName,
    newExpressionFormula,
    newExpressionScope,
    selectedWorkbookForExpression,
    workbookGridItems,
    engine,
    markUnsavedChanges,
  ]);

  // Delete named expression
  const deleteNamedExpression = useCallback(
    (name: string, isGlobal: boolean, workbookName?: string) => {
      try {
        const targetWorkbook = workbookName || workbookGridItems[0]?.name;
        const activeSheet = workbookGridItems.find(w => w.name === targetWorkbook)?.activeSheet;
        
        const success = engine.removeNamedExpression({
          expressionName: name,
          sheetName: isGlobal ? undefined : activeSheet,
          workbookName: isGlobal ? undefined : targetWorkbook,
        });

        if (success) {
          markUnsavedChanges();
        }
      } catch (error) {
        console.error("Failed to delete named expression:", error);
      }
    },
    [engine, workbookGridItems, markUnsavedChanges]
  );

  // Create WorkbookComponent for grid
  const WorkbookComponent = useCallback(({ workbookName }: { workbookName: string }) => {
    const workbookItem = workbookGridItems.find(item => item.name === workbookName);
    if (!workbookItem) return null;

    return (
      <div className="w-full h-full border border-gray-300 rounded-lg overflow-hidden bg-white">
        <div className="h-8 bg-gray-100 border-b border-gray-300 px-3 flex items-center justify-between">
          <span className="text-sm font-medium text-gray-700">{workbookName}</span>
          <span className="text-xs text-gray-500">Active: {workbookItem.activeSheet}</span>
        </div>
        <div className="flex-1 h-full">
          <SpreadsheetWithFormulaBar
            key={`${workbookName}-${workbookItem.activeSheet}`}
            sheetName={workbookItem.activeSheet}
            workbookName={workbookName}
            engine={engine}
            tables={tables}
            globalNamedExpressions={globalNamedExpressions}
            verboseErrors={verboseErrors}
          />
        </div>
      </div>
    );
  }, [workbookGridItems, engine, tables, globalNamedExpressions, verboseErrors, updateWorkbookActiveSheet]);

  // Create grid children from workbook items
  const gridChildren: GridChild[] = useMemo(() => {
    return workbookGridItems.map(item => ({
      id: item.name,
      title: item.name,
      x: item.x,
      y: item.y,
      width: item.width,
      height: item.height,
      component: "WorkbookComponent",
    }));
  }, [workbookGridItems]);

  // Grid components map
  const gridComponents = useMemo(() => ({
    WorkbookComponent: (gridChild: GridChild) => 
      <WorkbookComponent workbookName={gridChild.id} />,
  }), [WorkbookComponent]);

  return (
    <div className="h-full flex flex-col">
      {/* Grid-based header */}
      <div className="border-b border-gray-200 bg-gray-50">
        <div className="p-3 border-b border-gray-200">
          <div className="flex items-center justify-between">
            <h1 className="text-lg font-semibold text-gray-800">
              FormulaEngine Multi-Workbook Demo
            </h1>
            <div className="flex items-center gap-4 text-sm text-gray-600">
              <span data-testid="total-workbooks-count">
                Workbooks: {workbookGridItems.length}
              </span>
              <div className="flex items-center gap-2">
                <Button
                  size="sm"
                  variant="outline"
                  className="border-gray-300 text-gray-700"
                  onClick={addWorkbook}
                  data-testid="add-workbook-button"
                >
                  <Plus className="h-4 w-4 mr-1" />
                  Add Workbook
                </Button>
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
                  variant={verboseErrors ? "default" : "outline"}
                  className={`
                    ${
                      verboseErrors
                        ? "bg-red-600 hover:bg-red-700 text-white"
                        : "border-gray-300 text-gray-700"
                    }
                  `}
                  onClick={() => setVerboseErrors(!verboseErrors)}
                  data-testid="verbose-errors-toggle"
                >
                  <AlertTriangle className="h-4 w-4 mr-1" />
                  Verbose Errors
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
                  <span
                    className="text-xs text-orange-600 font-medium"
                    data-testid="unsaved-changes-indicator"
                  >
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
        <div
          className="border-b border-gray-200 bg-gray-50 p-4"
          data-testid="expressions-tables-panel"
        >
          <div className="space-y-4">
            {/* Add New Named Expression - Unified Form */}
            <div className="bg-white p-3 rounded border border-gray-200">
              <h3 className="text-sm font-semibold text-gray-800 mb-3">
                Add Named Expression
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">
                    Name
                  </label>
                  <Input
                    placeholder={
                      newExpressionScope === "sheet"
                        ? "e.g., COMMISSION"
                        : "e.g., TAX_RATE"
                    }
                    value={newExpressionName}
                    onChange={(e) => setNewExpressionName(e.target.value)}
                    className="text-sm"
                    data-testid="expression-name-input"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">
                    Formula
                  </label>
                  <Input
                    placeholder={
                      newExpressionScope === "sheet"
                        ? "e.g., 0.05"
                        : "e.g., 0.08"
                    }
                    value={newExpressionFormula}
                    onChange={(e) => setNewExpressionFormula(e.target.value)}
                    className="text-sm"
                    data-testid="expression-formula-input"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">
                    Scope
                  </label>
                  <div className="flex gap-2">
                    <select
                      value={newExpressionScope}
                      onChange={(e) =>
                        setNewExpressionScope(
                          e.target.value as "global" | "sheet"
                        )
                      }
                      className="flex-1 p-2 text-sm border border-gray-300 rounded-md"
                      data-testid="expression-scope-select"
                    >
                      <option value="global">Global</option>
                      <option value="sheet">Sheet</option>
                    </select>
                    {newExpressionScope === "sheet" && (
                      <select
                        value={selectedWorkbookForExpression}
                        onChange={(e) => setSelectedWorkbookForExpression(e.target.value)}
                        className="flex-1 p-2 text-sm border border-gray-300 rounded-md"
                        data-testid="workbook-select"
                      >
                        <option value="">Select Workbook</option>
                        {workbookGridItems.map(item => (
                          <option key={item.name} value={item.name}>{item.name}</option>
                        ))}
                      </select>
                    )}
                  </div>
                </div>
                <div className="flex items-end">
                  <Button
                    size="sm"
                    onClick={addNamedExpression}
                    disabled={
                      !newExpressionName.trim() || !newExpressionFormula.trim()
                    }
                    className="w-full"
                    data-testid="add-expression-button"
                  >
                    <Plus className="h-4 w-4 mr-1" />
                    Add
                  </Button>
                </div>
              </div>
            </div>

            {/* Existing Named Expressions and Tables */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* Global Named Expressions */}
              <div className="bg-white p-3 rounded border border-gray-200">
                <h3 className="text-sm font-semibold text-gray-800 mb-3">
                  Global Named Expressions ({globalNamedExpressions.size})
                </h3>
                <div className="space-y-2 max-h-40 overflow-y-auto">
                  {globalNamedExpressions.size === 0 ? (
                    <p className="text-xs text-gray-500 italic">
                      No global named expressions
                    </p>
                  ) : (
                    Array.from(globalNamedExpressions.entries()).map(
                      ([name, expr]) => (
                        <div
                          key={name}
                          className="flex items-center justify-between bg-gray-50 p-2 rounded"
                        >
                          <div className="flex-1 min-w-0">
                            <div className="font-medium text-xs text-gray-800 truncate">
                              {name}
                            </div>
                            <div className="text-xs text-gray-600 truncate">
                              {expr.expression}
                            </div>
                          </div>
                          <div className="flex items-center gap-1">
                            <Button
                              size="sm"
                              variant="ghost"
                              className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                              onClick={() =>
                                deleteNamedExpression(name, true)
                              }
                              data-testid={`delete-global-named-expression-${name}`}
                            >
                              <X className="h-3 w-3" />
                            </Button>
                          </div>
                        </div>
                      )
                    )
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
                      <div
                        key={name}
                        className="flex items-center justify-between bg-gray-50 p-2 rounded"
                      >
                        <div className="flex-1 min-w-0">
                          <div className="font-medium text-xs text-gray-800 truncate">
                            {name}
                          </div>
                          <div className="text-xs text-gray-600 truncate">
                            {table.sheetName} • {table.start.rowIndex + 1}:
                            {String.fromCharCode(65 + table.start.colIndex)}
                          </div>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Main infinite grid area */}
      <div className="flex-1 overflow-hidden">
        <Grid
          children={gridChildren}
          components={gridComponents}
          initialVisibleRect={{ x: 0, y: 0, w: 1200, h: 800 }}
          gridSize={20}
          style={{ width: "100%", height: "100%" }}
        />
      </div>

    </div>
  );
}
