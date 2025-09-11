import type { GridChild, ViewportState } from "@anocca-pub/components";
import { Grid } from "@anocca-pub/components";
import {
  AlertTriangle,
  Calculator,
  ChevronDown,
  ChevronUp,
  Edit2,
  Plus,
  Save,
  Trash2,
  X,
} from "lucide-react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { FormulaEngine } from "../src/core/engine";
import { useEngine } from "../src/react/hooks";
import { SpreadsheetWithFormulaBar } from "./components/SpreadsheetWithFormulaBar";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";

interface WorkbookGridItem {
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  activeSheet: string;
}

interface SavedState {
  workbookGridItems: WorkbookGridItem[];
  serializedEngine: string;
  viewport?: ViewportState;
}

const STORAGE_KEY = "formula-engine-excel-demo";

const loadFromLocalStorage = (): undefined | SavedState => {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return undefined;
  try {
    const { serializedEngine, workbookGridItems, viewport } = JSON.parse(saved);
    return { serializedEngine, workbookGridItems, viewport };
  } catch (error) {
    console.error("Failed to load from localStorage:", error);
    return undefined;
  }
};

const createEngine = () => {
  const engine = FormulaEngine.buildEmpty();
  const savedData = loadFromLocalStorage();

  if (savedData) {
    engine.resetToSerializedEngine(savedData.serializedEngine);

    return {
      engine,
      workbookGridItems: savedData.workbookGridItems,
      viewport: savedData.viewport,
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
    };

    return {
      engine,
      workbookGridItems: [defaultWorkbookItem],
      viewport: undefined,
    };
  }
};

const {
  engine,
  workbookGridItems: initialWorkbookGridItems,
  viewport: initialViewport,
} = createEngine();
export function ExcelDemo() {
  const [workbookGridItems, setWorkbookGridItems] = useState<
    WorkbookGridItem[]
  >(initialWorkbookGridItems);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [debugMode, setDebugMode] = useState(false);
  const [verboseErrors, setVerboseErrors] = useState(false);
  const [viewport, _setViewport] = useState<ViewportState | undefined>(
    initialViewport
  );

  const setViewport: typeof _setViewport = useCallback(
    (viewport) => {
      _setViewport(viewport);
      setHasUnsavedChanges(true);
    },
    [_setViewport]
  );

  // Named expressions UI state
  const [showNamedExpressions, setShowNamedExpressions] = useState(false);
  const [newExpressionName, setNewExpressionName] = useState("");
  const [newExpressionFormula, setNewExpressionFormula] = useState("");
  const [newExpressionScope, setNewExpressionScope] = useState<
    "global" | "workbook" | "sheet"
  >("global");

  // Sheet management state
  const [renamingSheet, setRenamingSheet] = useState<{
    workbookName: string;
    sheetName: string;
  } | null>(null);
  const [newSheetName, setNewSheetName] = useState("");

  const engineState = useEngine(engine);

  // Save to localStorage
  const saveToLocalStorage = () => {
    try {
      const dataToSave: SavedState = {
        workbookGridItems,
        serializedEngine: engine.serializeEngine(),
        viewport,
      };

      localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToSave));
      setHasUnsavedChanges(false);
      console.log("Spreadsheet saved to localStorage");
      console.log(
        "Viewport data saved:",
        workbookGridItems.map((item) => ({
          name: item.name,
          position: { x: item.x, y: item.y },
          size: { width: item.width, height: item.height },
        }))
      );
    } catch (error) {
      console.error("Failed to save to localStorage:", error);
    }
  };

  // Mark as having unsaved changes when sheets change
  const markUnsavedChanges = useCallback(() => {
    setHasUnsavedChanges(true);
  }, []);

  // Track changes
  useEffect(() => {
    const unsubscribe = engine.onUpdate(markUnsavedChanges);
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
    };

    setWorkbookGridItems((prev) => [...prev, newGridItem]);
    markUnsavedChanges();
  }, [workbookGridItems.length, engine, markUnsavedChanges]);

  // Update active sheet for a workbook
  const updateWorkbookActiveSheet = useCallback(
    (workbookName: string, sheetName: string) => {
      setWorkbookGridItems((prev) =>
        prev.map((item) =>
          item.name === workbookName
            ? { ...item, activeSheet: sheetName }
            : item
        )
      );
      markUnsavedChanges();
    },
    [markUnsavedChanges]
  );

  // Add named expression
  const addNamedExpression = useCallback(() => {
    if (newExpressionName.trim() && newExpressionFormula.trim()) {
      // Use the first workbook as the active workbook for scoped expressions
      const workbookName = workbookGridItems[0]?.name;
      const activeSheet = workbookGridItems[0]?.activeSheet;

      engine.addNamedExpression({
        expressionName: newExpressionName.trim(),
        expression: newExpressionFormula.trim(),
        sheetName: newExpressionScope === "sheet" ? activeSheet : undefined,
        workbookName:
          newExpressionScope === "workbook" || newExpressionScope === "sheet"
            ? workbookName
            : undefined,
      });

      setNewExpressionName("");
      setNewExpressionFormula("");
      markUnsavedChanges();
    }
  }, [
    newExpressionName,
    newExpressionFormula,
    newExpressionScope,
    workbookGridItems,
    engine,
    markUnsavedChanges,
  ]);

  // Delete named expression
  const deleteNamedExpression = useCallback(
    (name: string, isGlobal: boolean, workbookName?: string) => {
      try {
        const targetWorkbook = workbookName || workbookGridItems[0]?.name;
        const activeSheet = workbookGridItems.find(
          (w) => w.name === targetWorkbook
        )?.activeSheet;

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

  // Rename sheet
  const renameSheet = useCallback(
    (workbookName: string, oldSheetName: string, newSheetName: string) => {
      try {
        engine.renameSheet({
          workbookName,
          sheetName: oldSheetName,
          newSheetName: newSheetName.trim(),
        });

        // Update active sheet if it was the renamed one
        updateWorkbookActiveSheet(workbookName, newSheetName.trim());
        markUnsavedChanges();
        setRenamingSheet(null);
        setNewSheetName("");
      } catch (error) {
        console.error("Failed to rename sheet:", error);
      }
    },
    [engine, updateWorkbookActiveSheet, markUnsavedChanges]
  );

  // Delete sheet
  const deleteSheet = useCallback(
    (workbookName: string, sheetName: string) => {
      try {
        const sheets = engine.getSheets(workbookName);
        if (sheets.size <= 1) {
          alert("Cannot delete the last sheet in a workbook");
          return;
        }

        engine.removeSheet({ workbookName, sheetName });

        // If we deleted the active sheet, switch to the first available sheet
        const workbookItem = workbookGridItems.find(
          (item) => item.name === workbookName
        );
        if (workbookItem?.activeSheet === sheetName) {
          const remainingSheets = engine.getSheets(workbookName);
          const firstSheet = Array.from(remainingSheets.keys())[0];
          if (firstSheet) {
            updateWorkbookActiveSheet(workbookName, firstSheet);
          }
        }

        markUnsavedChanges();
      } catch (error) {
        console.error("Failed to delete sheet:", error);
      }
    },
    [engine, workbookGridItems, updateWorkbookActiveSheet, markUnsavedChanges]
  );

  // Start renaming sheet
  const startRenaming = useCallback(
    (workbookName: string, sheetName: string) => {
      setRenamingSheet({ workbookName, sheetName });
      setNewSheetName(sheetName);
    },
    []
  );

  // Cancel renaming
  const cancelRenaming = useCallback(() => {
    setRenamingSheet(null);
    setNewSheetName("");
  }, []);

  // Create WorkbookComponent for grid
  const WorkbookComponent = useCallback(
    ({ workbookName }: { workbookName: string }) => {
      const workbookItem = workbookGridItems.find(
        (item) => item.name === workbookName
      );
      if (!workbookItem) return null;

      // Get all sheets for this workbook
      const sheets = engine.getSheets(workbookName);
      const sheetNames = Array.from(sheets.keys());

      // Add new sheet handler
      const addSheet = () => {
        const newSheetCount = sheetNames.length + 1;
        const newSheetName = `Sheet${newSheetCount}`;

        try {
          const result = engine.addSheet({
            workbookName,
            sheetName: newSheetName,
          });

          // Switch to the new sheet
          updateWorkbookActiveSheet(workbookName, result.name);
          markUnsavedChanges();
        } catch (error) {
          console.error("Failed to add sheet:", error);
        }
      };

      return (
        <div className="w-full h-full border border-gray-300 overflow-hidden bg-white flex flex-col">
          {/* Workbook Header */}
          <div className="h-8 bg-gray-100 border-b border-gray-300 px-3 flex items-center justify-between flex-shrink-0">
            <span className="text-sm font-medium text-gray-700">
              {workbookName}
            </span>
            <span
              className="text-xs text-gray-500"
              data-testid={`sheet-count-${workbookName}`}
            >
              Sheets: {sheetNames.length}
            </span>
          </div>

          {/* Spreadsheet Content */}
          <div className="flex-1 overflow-hidden">
            <SpreadsheetWithFormulaBar
              key={`${workbookName}-${workbookItem.activeSheet}`}
              sheetName={workbookItem.activeSheet}
              workbookName={workbookName}
              engine={engine}
              verboseErrors={verboseErrors}
            />
          </div>

          {/* Sheet Tabs at Bottom (Excel-style) */}
          <div className="h-8 bg-gray-50 border-t border-gray-200 flex items-center px-2 flex-shrink-0">
            <div className="flex items-center gap-1">
              {sheetNames.map((sheetName) => {
                const isRenaming =
                  renamingSheet?.workbookName === workbookName &&
                  renamingSheet?.sheetName === sheetName;

                return (
                  <div key={sheetName} className="flex items-center group">
                    {isRenaming ? (
                      <div className="flex items-center gap-1 bg-white border border-blue-500 rounded px-2 py-1">
                        <input
                          type="text"
                          value={newSheetName}
                          onChange={(e) => setNewSheetName(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") {
                              renameSheet(
                                workbookName,
                                sheetName,
                                newSheetName
                              );
                            } else if (e.key === "Escape") {
                              cancelRenaming();
                            }
                          }}
                          onBlur={() => {
                            if (
                              newSheetName.trim() &&
                              newSheetName.trim() !== sheetName
                            ) {
                              renameSheet(
                                workbookName,
                                sheetName,
                                newSheetName
                              );
                            } else {
                              cancelRenaming();
                            }
                          }}
                          className="text-xs w-20 outline-none bg-transparent"
                          autoFocus
                          data-testid={`rename-sheet-input-${sheetName}`}
                        />
                      </div>
                    ) : (
                      <div
                        className={`
                          px-3 py-1 text-xs font-medium rounded-t border-b-2 whitespace-nowrap relative flex items-center gap-1 group cursor-pointer
                          ${
                            workbookItem.activeSheet === sheetName
                              ? "bg-white text-blue-600 border-blue-500 -mb-px"
                              : "bg-transparent text-gray-600 border-transparent hover:bg-gray-100"
                          }
                        `}
                        onClick={() =>
                          updateWorkbookActiveSheet(workbookName, sheetName)
                        }
                        onDoubleClick={() =>
                          startRenaming(workbookName, sheetName)
                        }
                        data-testid={`sheet-tab-${sheetName}`}
                      >
                        <span>{sheetName}</span>

                        {/* Sheet actions (visible on hover, inside tab) */}
                        <div className="opacity-0 group-hover:opacity-100 transition-opacity flex items-center gap-1 ml-1">
                          <div
                            className="h-4 w-4 p-0 text-gray-500 hover:text-blue-600 flex items-center justify-center cursor-pointer"
                            onClick={(e) => {
                              e.stopPropagation();
                              startRenaming(workbookName, sheetName);
                            }}
                            title="Rename Sheet"
                            data-testid={`rename-sheet-${sheetName}`}
                          >
                            <Edit2 className="h-3 w-3" />
                          </div>
                          {sheetNames.length > 1 && (
                            <div
                              className="h-4 w-4 p-0 text-gray-500 hover:text-red-600 flex items-center justify-center cursor-pointer"
                              onClick={(e) => {
                                e.stopPropagation();
                                if (
                                  confirm(
                                    `Are you sure you want to delete sheet "${sheetName}"?`
                                  )
                                ) {
                                  deleteSheet(workbookName, sheetName);
                                }
                              }}
                              title="Delete Sheet"
                              data-testid={`delete-sheet-${sheetName}`}
                            >
                              <Trash2 className="h-3 w-3" />
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Add Sheet Button (Excel-style, to the right) */}
            <div className="flex items-center ml-2">
              <Button
                size="sm"
                variant="ghost"
                className="h-6 w-6 p-0 text-gray-600 hover:text-gray-800 border border-gray-300 rounded"
                onClick={addSheet}
                title="Add Sheet"
                data-testid={`add-sheet-${workbookName}`}
              >
                <Plus className="h-3 w-3" />
              </Button>
            </div>
          </div>
        </div>
      );
    },
    [
      workbookGridItems,
      engine,
      verboseErrors,
      updateWorkbookActiveSheet,
      markUnsavedChanges,
      renamingSheet,
      newSheetName,
      renameSheet,
      cancelRenaming,
      startRenaming,
      deleteSheet,
    ]
  );

  // Create grid children from workbook items
  const gridChildren: GridChild[] = useMemo(() => {
    return workbookGridItems.map((item) => ({
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
  const gridComponents = useMemo(
    () => ({
      WorkbookComponent: (gridChild: GridChild) => (
        <WorkbookComponent workbookName={gridChild.id} />
      ),
    }),
    [WorkbookComponent]
  );

  const numTables = engineState.tables
    .values()
    .reduce((sum, wb) => sum + wb.size, 0);

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
              {workbookGridItems.length > 0 && (
                <span className="text-blue-600 font-medium">
                  Active:{" "}
                  {workbookGridItems.find((w) => w.name === "Workbook1")
                    ?.name || workbookGridItems[0]?.name}{" "}
                  →{" "}
                  {workbookGridItems.find((w) => w.name === "Workbook1")
                    ?.activeSheet || workbookGridItems[0]?.activeSheet}
                </span>
              )}
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
                    placeholder={"e.g., COMMISSION"}
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
                    placeholder={"e.g., 0.05"}
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
                  <select
                    value={newExpressionScope}
                    onChange={(e) =>
                      setNewExpressionScope(
                        e.target.value as "global" | "workbook" | "sheet"
                      )
                    }
                    className="w-full p-2 text-sm border border-gray-300 rounded-md"
                    data-testid="expression-scope-select"
                  >
                    <option value="global">Global</option>
                    <option value="workbook">
                      Workbook ({workbookGridItems[0]?.name || "None"})
                    </option>
                    <option value="sheet">
                      Sheet ({workbookGridItems[0]?.name || "None"} →{" "}
                      {workbookGridItems[0]?.activeSheet || "None"})
                    </option>
                  </select>
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
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {/* All Named Expressions - Consolidated */}
              <div
                className="bg-white p-3 rounded border border-gray-200"
                data-testid="named-expressions-section"
              >
                <h3
                  className="text-sm font-semibold text-gray-800 mb-3"
                  data-testid="named-expressions-title"
                >
                  Named Expressions (
                  {engineState.namedExpressions.globalExpressions.size +
                    Array.from(
                      engineState.namedExpressions.workbookExpressions.values()
                    ).reduce((sum, wb) => sum + wb.size, 0) +
                    Array.from(
                      engineState.namedExpressions.sheetExpressions.values()
                    ).reduce(
                      (sum, wb) =>
                        sum +
                        Array.from(wb.values()).reduce(
                          (sheetSum, sheet) => sheetSum + sheet.size,
                          0
                        ),
                      0
                    )}
                  )
                </h3>
                <div
                  className="space-y-2 max-h-40 overflow-y-auto"
                  data-testid="named-expressions-list"
                >
                  {/* Global expressions */}
                  {Array.from(
                    engineState.namedExpressions.globalExpressions.entries()
                  ).map(([name, expr]) => (
                    <div
                      key={`global-${name}`}
                      className="flex items-center justify-between bg-gray-50 p-2 rounded"
                      data-testid={`named-expression-${name}`}
                    >
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2 mb-1">
                          <div
                            className="font-medium text-xs text-gray-800 truncate"
                            data-testid={`expression-name-${name}`}
                          >
                            {name}
                          </div>
                          <span
                            className="px-1.5 py-0.5 text-xs bg-gray-200 text-gray-700 rounded"
                            data-testid={`expression-scope-${name}`}
                          >
                            Global
                          </span>
                        </div>
                        <div
                          className="text-xs text-gray-600 truncate"
                          data-testid={`expression-formula-${name}`}
                        >
                          {expr.expression}
                        </div>
                      </div>
                      <div className="flex items-center gap-1">
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
                    </div>
                  ))}

                  {/* Workbook-scoped expressions */}
                  {Array.from(
                    engineState.namedExpressions.workbookExpressions.entries()
                  ).flatMap(([workbookName, expressions]) =>
                    Array.from(expressions.entries()).map(([name, expr]) => (
                      <div
                        key={`workbook-${workbookName}-${name}`}
                        className="flex items-center justify-between bg-blue-50 p-2 rounded"
                        data-testid={`named-expression-${name}`}
                      >
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 mb-1">
                            <div
                              className="font-medium text-xs text-gray-800 truncate"
                              data-testid={`expression-name-${name}`}
                            >
                              {name}
                            </div>
                            <span
                              className="px-1.5 py-0.5 text-xs bg-blue-200 text-blue-800 rounded"
                              data-testid={`expression-scope-${name}`}
                            >
                              {workbookName}
                            </span>
                          </div>
                          <div
                            className="text-xs text-blue-600 truncate"
                            data-testid={`expression-formula-${name}`}
                          >
                            {expr.expression}
                          </div>
                        </div>
                        <div className="flex items-center gap-1">
                          <Button
                            size="sm"
                            variant="ghost"
                            className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                            onClick={() =>
                              deleteNamedExpression(name, false, workbookName)
                            }
                            data-testid={`delete-workbook-named-expression-${name}`}
                          >
                            <X className="h-3 w-3" />
                          </Button>
                        </div>
                      </div>
                    ))
                  )}

                  {/* Sheet-scoped expressions */}
                  {Array.from(
                    engineState.namedExpressions.sheetExpressions.entries()
                  ).flatMap(([workbookName, workbookSheets]) =>
                    Array.from(workbookSheets.entries()).flatMap(
                      ([sheetName, expressions]) =>
                        Array.from(expressions.entries()).map(
                          ([name, expr]) => (
                            <div
                              key={`sheet-${workbookName}-${sheetName}-${name}`}
                              className="flex items-center justify-between bg-green-50 p-2 rounded"
                              data-testid={`named-expression-${name}`}
                            >
                              <div className="flex-1 min-w-0">
                                <div className="flex items-center gap-2 mb-1">
                                  <div
                                    className="font-medium text-xs text-gray-800 truncate"
                                    data-testid={`expression-name-${name}`}
                                  >
                                    {name}
                                  </div>
                                  <span
                                    className="px-1.5 py-0.5 text-xs bg-green-200 text-green-800 rounded"
                                    data-testid={`expression-scope-${name}`}
                                  >
                                    {workbookName} → {sheetName}
                                  </span>
                                </div>
                                <div
                                  className="text-xs text-green-600 truncate"
                                  data-testid={`expression-formula-${name}`}
                                >
                                  {expr.expression}
                                </div>
                              </div>
                              <div className="flex items-center gap-1">
                                <Button
                                  size="sm"
                                  variant="ghost"
                                  className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                                  onClick={() =>
                                    deleteNamedExpression(
                                      name,
                                      false,
                                      workbookName
                                    )
                                  }
                                  data-testid={`delete-sheet-named-expression-${name}`}
                                >
                                  <X className="h-3 w-3" />
                                </Button>
                              </div>
                            </div>
                          )
                        )
                    )
                  )}

                  {/* Show message if no expressions */}
                  {engineState.namedExpressions.globalExpressions.size === 0 &&
                    engineState.namedExpressions.workbookExpressions.size ===
                      0 &&
                    engineState.namedExpressions.sheetExpressions.size ===
                      0 && (
                      <p className="text-xs text-gray-500 italic">
                        No named expressions
                      </p>
                    )}
                </div>
              </div>

              {/* Tables */}
              <div
                className="bg-white p-3 rounded border border-gray-200"
                data-testid="tables-section"
              >
                <h3
                  className="text-sm font-semibold text-gray-800 mb-3"
                  data-testid="tables-title"
                >
                  Tables ({numTables})
                </h3>
                <div
                  className="space-y-2 max-h-40 overflow-y-auto"
                  data-testid="tables-list"
                >
                  {numTables === 0 ? (
                    <p className="text-xs text-gray-500 italic">No tables</p>
                  ) : (
                    Array.from(engineState.tables).flatMap(
                      ([workbookName, workbook]) =>
                        Array.from(workbook.entries()).map(
                          ([tableName, table]) => (
                            <div
                              key={tableName + "|" + workbookName}
                              className="flex items-center justify-between bg-gray-50 p-2 rounded"
                              data-testid={`table-${tableName}`}
                            >
                              <div className="flex-1 min-w-0">
                                <div
                                  className="font-medium text-xs text-gray-800 truncate"
                                  data-testid={`table-name-${tableName}`}
                                >
                                  {tableName}
                                </div>
                                <div className="text-xs text-gray-600 truncate">
                                  {table.sheetName} • {workbookName}
                                </div>
                                <div className="text-xs text-gray-500">
                                  Range:{" "}
                                  {String.fromCharCode(
                                    65 + table.start.colIndex
                                  )}
                                  {table.start.rowIndex + 1}:
                                  {String.fromCharCode(
                                    65 +
                                      table.start.colIndex +
                                      table.headers.size
                                  )}
                                  {table.endRow.type === "number"
                                    ? table.endRow.value + 1
                                    : "∞"}
                                </div>
                              </div>
                            </div>
                          )
                        )
                    )
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
          viewport={viewport}
          onViewportChange={setViewport}
          gridSize={20}
          style={{ width: "100%", height: "100%" }}
          onChildrenChange={(newChildren) => {
            // Update workbook grid items when children change (move/resize)
            const updatedItems = workbookGridItems.map((item) => {
              const updatedChild = newChildren.find(
                (child) => child.id === item.name
              );
              if (updatedChild) {
                return {
                  ...item,
                  x: updatedChild.x,
                  y: updatedChild.y,
                  width: updatedChild.width,
                  height: updatedChild.height,
                };
              }
              return item;
            });
            setWorkbookGridItems(updatedItems);
            markUnsavedChanges();
          }}
        />
      </div>
    </div>
  );
}
