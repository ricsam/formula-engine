import type { GridChild, ViewportState } from "@anocca-pub/components";
import { Grid } from "@anocca-pub/components";
import {
  AlertTriangle,
  Calculator,
  ChevronDown,
  ChevronUp,
  Copy,
  Download,
  Edit2,
  File,
  Files,
  FileText,
  FolderOpen,
  Plus,
  Save,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { FormulaEngine } from "../src/core/engine";
import { useEngine } from "../src/react/hooks";
import { serialize, deserialize } from "../src/core/map-serializer";
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
  engineState: ReturnType<FormulaEngine["getState"]>;
  viewport?: ViewportState;
}

// OPFS (Origin Private File System) helper functions
const OPFS_DIRECTORY = "formula-engine";
const DEFAULT_FILE = "workbook.json";

async function getOPFSRoot(): Promise<FileSystemDirectoryHandle> {
  return await navigator.storage.getDirectory();
}

async function getOPFSDirectory(): Promise<FileSystemDirectoryHandle> {
  const root = await getOPFSRoot();
  return await root.getDirectoryHandle(OPFS_DIRECTORY, { create: true });
}

async function listOPFSFiles(): Promise<string[]> {
  try {
    const dir = await getOPFSDirectory();
    const files: string[] = [];
    // @ts-ignore - AsyncIterator not fully typed
    for await (const [name, handle] of dir.entries()) {
      if (handle.kind === "file" && name.endsWith(".json")) {
        files.push(name);
      }
    }
    return files.sort();
  } catch (error) {
    console.error("Failed to list OPFS files:", error);
    return [];
  }
}

async function saveToOPFS(filename: string, data: SavedState): Promise<void> {
  const dir = await getOPFSDirectory();
  const fileHandle = await dir.getFileHandle(filename, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(serialize(data));
  await writable.close();
}

async function loadFromOPFS(filename: string): Promise<SavedState | null> {
  try {
    const dir = await getOPFSDirectory();
    const fileHandle = await dir.getFileHandle(filename);
    const file = await fileHandle.getFile();
    const contents = await file.text();
    return deserialize(contents) as SavedState;
  } catch (error) {
    console.error(`Failed to load ${filename} from OPFS:`, error);
    return null;
  }
}

async function deleteFromOPFS(filename: string): Promise<void> {
  try {
    const dir = await getOPFSDirectory();
    await dir.removeEntry(filename);
  } catch (error) {
    console.error(`Failed to delete ${filename} from OPFS:`, error);
    throw error;
  }
}

async function renameInOPFS(oldName: string, newName: string): Promise<void> {
  try {
    const dir = await getOPFSDirectory();
    const oldHandle = await dir.getFileHandle(oldName);
    const file = await oldHandle.getFile();
    const contents = await file.text();
    
    // Create new file with new name
    const newHandle = await dir.getFileHandle(newName, { create: true });
    const writable = await newHandle.createWritable();
    await writable.write(contents);
    await writable.close();
    
    // Delete old file
    await dir.removeEntry(oldName);
  } catch (error) {
    console.error(`Failed to rename ${oldName} to ${newName}:`, error);
    throw error;
  }
}

const createEngine = () => {
  const engine = FormulaEngine.buildEmpty();
  
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
  const [currentFileName, setCurrentFileName] = useState<string>(DEFAULT_FILE);
  const [opfsFiles, setOpfsFiles] = useState<string[]>([]);
  const [showFileManager, setShowFileManager] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

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

  // Workbook management state
  const [renamingWorkbook, setRenamingWorkbook] = useState<string | null>(null);
  const [newWorkbookName, setNewWorkbookName] = useState("");

  const engineState = useEngine(engine);

  // Load from OPFS on mount
  useEffect(() => {
    const loadInitialFile = async () => {
      try {
        const files = await listOPFSFiles();
        setOpfsFiles(files);

        // Try to load the default file or the first available file
        const fileToLoad = files.includes(DEFAULT_FILE) ? DEFAULT_FILE : files[0];
        
        if (fileToLoad) {
          const data = await loadFromOPFS(fileToLoad);
          if (data && data.engineState && data.workbookGridItems) {
            // Reset the engine using the loaded state
            engine._workbookManager.resetWorkbooks(data.engineState.workbooks);

            data.engineState.workbooks.forEach((workbook) => {
              engine._namedExpressionManager.addWorkbook(workbook.name);
              engine._tableManager.addWorkbook(workbook.name);
              workbook.sheets.forEach((sheet) => {
                engine._namedExpressionManager.addSheet({
                  workbookName: workbook.name,
                  sheetName: sheet.name,
                });
              });
            });

            engine._namedExpressionManager.resetNamedExpressions(
              data.engineState.namedExpressions
            );
            engine._tableManager.resetTables(data.engineState.tables);

            engine.reevaluate();
            engine._eventManager.emitUpdate();

            setWorkbookGridItems(data.workbookGridItems);
            if (data.viewport) {
              _setViewport(data.viewport);
            }
            setCurrentFileName(fileToLoad);
            setHasUnsavedChanges(false);
            console.log(`Loaded ${fileToLoad} from OPFS`);
          }
        }
      } catch (error) {
        console.error("Failed to load from OPFS:", error);
      } finally {
        setIsLoading(false);
      }
    };

    loadInitialFile();
  }, [engine, _setViewport]);

  // Auto-save to OPFS
  const saveToOPFSAuto = useCallback(async () => {
    try {
      const dataToSave: SavedState = {
        workbookGridItems,
        engineState: engine.getState(),
        viewport,
      };

      await saveToOPFS(currentFileName, dataToSave);
      setHasUnsavedChanges(false);
      console.log(`Auto-saved to OPFS: ${currentFileName}`);
      
      // Refresh file list
      const files = await listOPFSFiles();
      setOpfsFiles(files);
    } catch (error) {
      console.error("Failed to auto-save to OPFS:", error);
    }
  }, [workbookGridItems, engine, viewport, currentFileName]);

  // Manual save (for immediate feedback)
  const saveFile = useCallback(async () => {
    await saveToOPFSAuto();
  }, [saveToOPFSAuto]);

  // Load a file from OPFS
  const loadFile = useCallback(async (filename: string) => {
    try {
      const data = await loadFromOPFS(filename);
      if (!data || !data.engineState || !data.workbookGridItems) {
        throw new Error("Invalid file format");
      }

      // Reset the engine using the loaded state
      engine._workbookManager.resetWorkbooks(data.engineState.workbooks);

      data.engineState.workbooks.forEach((workbook) => {
        engine._namedExpressionManager.addWorkbook(workbook.name);
        engine._tableManager.addWorkbook(workbook.name);
        workbook.sheets.forEach((sheet) => {
          engine._namedExpressionManager.addSheet({
            workbookName: workbook.name,
            sheetName: sheet.name,
          });
        });
      });

      engine._namedExpressionManager.resetNamedExpressions(
        data.engineState.namedExpressions
      );
      engine._tableManager.resetTables(data.engineState.tables);

      engine.reevaluate();
      engine._eventManager.emitUpdate();

      setWorkbookGridItems(data.workbookGridItems);
      if (data.viewport) {
        _setViewport(data.viewport);
      }
      setCurrentFileName(filename);
      setHasUnsavedChanges(false);
      console.log(`Loaded ${filename} from OPFS`);
    } catch (error) {
      console.error(`Failed to load ${filename}:`, error);
      alert(`Failed to load file: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }, [engine, _setViewport]);

  // Create new file
  const newFile = useCallback(async () => {
    if (hasUnsavedChanges) {
      if (!confirm("You have unsaved changes. Are you sure you want to create a new file?")) {
        return;
      }
    }

    // Prompt for new filename
    const newFileName = prompt("Enter new file name (without .json extension):", "workbook");
    if (!newFileName) return;

    const fullFileName = newFileName.endsWith(".json") ? newFileName : `${newFileName}.json`;

    // Check if file already exists
    if (opfsFiles.includes(fullFileName)) {
      if (!confirm(`File "${fullFileName}" already exists. Overwrite?`)) {
        return;
      }
    }

    // Reset to initial state
    engine._workbookManager.resetWorkbooks(new Map());
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

    setWorkbookGridItems([defaultWorkbookItem]);
    _setViewport(undefined);
    setCurrentFileName(fullFileName);
    setHasUnsavedChanges(true);
    
    engine.reevaluate();
    engine._eventManager.emitUpdate();

    // Save the new file immediately
    const dataToSave: SavedState = {
      workbookGridItems: [defaultWorkbookItem],
      engineState: engine.getState(),
      viewport: undefined,
    };
    await saveToOPFS(fullFileName, dataToSave);
    setHasUnsavedChanges(false);

    // Refresh file list
    const files = await listOPFSFiles();
    setOpfsFiles(files);
  }, [engine, hasUnsavedChanges, _setViewport, opfsFiles]);

  // Delete a file from OPFS
  const deleteFile = useCallback(async (filename: string) => {
    if (!confirm(`Are you sure you want to delete "${filename}"?`)) {
      return;
    }

    try {
      await deleteFromOPFS(filename);
      console.log(`Deleted ${filename} from OPFS`);

      // Refresh file list
      const files = await listOPFSFiles();
      setOpfsFiles(files);

      // If we deleted the current file, create a new one
      if (filename === currentFileName) {
        await newFile();
      }
    } catch (error) {
      console.error(`Failed to delete ${filename}:`, error);
      alert(`Failed to delete file: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }, [currentFileName, newFile]);

  // Rename current file
  const renameFile = useCallback(async () => {
    const newName = prompt("Enter new file name (without .json extension):", currentFileName.replace(".json", ""));
    if (!newName) return;

    const fullNewName = newName.endsWith(".json") ? newName : `${newName}.json`;
    
    if (opfsFiles.includes(fullNewName)) {
      alert(`File "${fullNewName}" already exists.`);
      return;
    }

    try {
      await renameInOPFS(currentFileName, fullNewName);
      setCurrentFileName(fullNewName);
      
      // Refresh file list
      const files = await listOPFSFiles();
      setOpfsFiles(files);
      
      console.log(`Renamed ${currentFileName} to ${fullNewName}`);
    } catch (error) {
      console.error("Failed to rename file:", error);
      alert(`Failed to rename file: ${error instanceof Error ? error.message : "Unknown error"}`);
    }
  }, [currentFileName, opfsFiles]);

  // Export current file to file system
  const exportToFileSystem = useCallback(() => {
    try {
      const dataToExport: SavedState = {
        workbookGridItems,
        engineState: engine.getState(),
        viewport,
      };

      const jsonString = serialize(dataToExport);
      const blob = new Blob([jsonString], { type: "application/json" });
      const url = URL.createObjectURL(blob);

      const link = document.createElement("a");
      link.href = url;
      link.download = currentFileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      console.log(`Exported ${currentFileName} to file system`);
    } catch (error) {
      console.error("Failed to export to file system:", error);
      alert("Failed to export. Please try again.");
    }
  }, [workbookGridItems, engine, viewport, currentFileName]);

  // Import file from file system to OPFS
  const importFromFileSystem = useCallback(async () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".json";

    input.onchange = async (event) => {
      const file = (event.target as HTMLInputElement).files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const jsonString = e.target?.result as string;
          const importedData: SavedState = deserialize(
            jsonString
          ) as SavedState;

          // Validate the imported data structure
          if (!importedData.engineState || !importedData.workbookGridItems) {
            throw new Error("Invalid file format: missing required fields");
          }

          // Save to OPFS with the imported filename
          let filename = file.name;
          if (opfsFiles.includes(filename)) {
            if (!confirm(`File "${filename}" already exists in OPFS. Overwrite?`)) {
              // Generate a unique name
              const baseName = filename.replace(".json", "");
              let counter = 1;
              while (opfsFiles.includes(`${baseName}-${counter}.json`)) {
                counter++;
              }
              filename = `${baseName}-${counter}.json`;
            }
          }

          await saveToOPFS(filename, importedData);

          // Load the imported file
          await loadFile(filename);

          // Refresh file list
          const files = await listOPFSFiles();
          setOpfsFiles(files);

          console.log(`Imported ${file.name} to OPFS as ${filename}`);
          alert(`File imported successfully as "${filename}"`);
        } catch (error) {
          console.error("Failed to import file:", error);
          alert(
            `Failed to import file: ${error instanceof Error ? error.message : "Unknown error"}`
          );
        }
      };

      reader.readAsText(file);
    };

    input.click();
  }, [engine, _setViewport, opfsFiles, loadFile]);

  // Mark as having unsaved changes when sheets change
  const markUnsavedChanges = useCallback(() => {
    setHasUnsavedChanges(true);
  }, []);

  // Track changes
  useEffect(() => {
    const unsubscribe = engine.onUpdate(markUnsavedChanges);
    return unsubscribe;
  }, [engine, markUnsavedChanges]);

  // Auto-save effect (debounced)
  useEffect(() => {
    if (!hasUnsavedChanges || isLoading) return;

    const timeoutId = setTimeout(() => {
      saveToOPFSAuto();
    }, 2000); // Auto-save after 2 seconds of inactivity

    return () => clearTimeout(timeoutId);
  }, [hasUnsavedChanges, saveToOPFSAuto, isLoading]);

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
      setWorkbookGridItems((prev) => {
        // Find the workbook item
        const workbookItem = prev.find(item => item.name === workbookName);
        
        // If workbook not found or sheet is already active, no update needed
        if (!workbookItem || workbookItem.activeSheet === sheetName) {
          return prev;
        }
        
        // Only update if there's an actual change
        return prev.map((item) =>
          item.name === workbookName
            ? { ...item, activeSheet: sheetName }
            : item
        );
      });
      
      // Only mark as unsaved if we found the workbook and it's a different sheet
      const workbookItem = workbookGridItems.find(item => item.name === workbookName);
      if (workbookItem && workbookItem.activeSheet !== sheetName) {
        markUnsavedChanges();
      }
    },
    [markUnsavedChanges, workbookGridItems]
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

  // Rename workbook
  const renameWorkbook = useCallback(
    (oldWorkbookName: string, newWorkbookName: string) => {
      try {
        engine.renameWorkbook({
          workbookName: oldWorkbookName,
          newWorkbookName: newWorkbookName.trim(),
        });

        // Update workbook grid items
        setWorkbookGridItems((prev) =>
          prev.map((item) =>
            item.name === oldWorkbookName
              ? { ...item, name: newWorkbookName.trim() }
              : item
          )
        );

        markUnsavedChanges();
        setRenamingWorkbook(null);
        setNewWorkbookName("");
      } catch (error) {
        console.error("Failed to rename workbook:", error);
      }
    },
    [engine, markUnsavedChanges]
  );

  // Start renaming workbook
  const startRenamingWorkbook = useCallback((workbookName: string) => {
    setRenamingWorkbook(workbookName);
    setNewWorkbookName(workbookName);
  }, []);

  // Cancel workbook renaming
  const cancelWorkbookRenaming = useCallback(() => {
    setRenamingWorkbook(null);
    setNewWorkbookName("");
  }, []);

  // Clone workbook
  const cloneWorkbook = useCallback(
    (sourceWorkbookName: string) => {
      try {
        // Generate a unique name for the cloned workbook
        let cloneNumber = 2;
        let cloneName = `${sourceWorkbookName} (Copy)`;

        // Keep incrementing until we find a unique name
        while (workbookGridItems.some((item) => item.name === cloneName)) {
          cloneName = `${sourceWorkbookName} (Copy ${cloneNumber})`;
          cloneNumber++;
        }

        // Clone the workbook in the engine
        engine.cloneWorkbook(sourceWorkbookName, cloneName);

        // Find the source workbook's grid item to position the clone nearby
        const sourceItem = workbookGridItems.find(
          (item) => item.name === sourceWorkbookName
        );
        const sourceX = sourceItem?.x || 100;
        const sourceY = sourceItem?.y || 100;

        // Create grid item for the cloned workbook
        const clonedGridItem: WorkbookGridItem = {
          name: cloneName,
          x: sourceX + 50, // Offset slightly from the original
          y: sourceY + 50,
          width: sourceItem?.width || 800,
          height: sourceItem?.height || 600,
          activeSheet: sourceItem?.activeSheet || "Sheet1",
        };

        setWorkbookGridItems((prev) => [...prev, clonedGridItem]);
        markUnsavedChanges();

        console.log(
          `Workbook "${sourceWorkbookName}" cloned as "${cloneName}"`
        );
      } catch (error) {
        console.error("Failed to clone workbook:", error);
        alert(
          `Failed to clone workbook: ${error instanceof Error ? error.message : "Unknown error"}`
        );
      }
    },
    [workbookGridItems, engine, markUnsavedChanges]
  );

  // Delete workbook
  const deleteWorkbook = useCallback(
    (workbookName: string) => {
      try {
        // Prevent deleting the last workbook
        if (workbookGridItems.length <= 1) {
          alert(
            "Cannot delete the last workbook. At least one workbook must remain."
          );
          return;
        }

        // Show confirmation dialog
        const confirmMessage = `Are you sure you want to delete workbook "${workbookName}"?\n\nThis will permanently delete all sheets, data, named expressions, and tables in this workbook. This action cannot be undone.`;

        if (!confirm(confirmMessage)) {
          return;
        }

        // Remove from engine
        engine.removeWorkbook(workbookName);

        // Remove from grid items
        setWorkbookGridItems((prev) =>
          prev.filter((item) => item.name !== workbookName)
        );

        markUnsavedChanges();
        console.log(`Workbook "${workbookName}" deleted successfully`);
      } catch (error) {
        console.error("Failed to delete workbook:", error);
        alert(
          `Failed to delete workbook: ${error instanceof Error ? error.message : "Unknown error"}`
        );
      }
    },
    [workbookGridItems, engine, markUnsavedChanges]
  );

  // Delete table
  const deleteTable = useCallback(
    (tableName: string, workbookName: string) => {
      try {
        const success = engine.removeTable({ tableName, workbookName });
        if (success) {
          markUnsavedChanges();
        }
      } catch (error) {
        console.error("Failed to delete table:", error);
      }
    },
    [engine, markUnsavedChanges]
  );

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
          <div className="h-8 bg-gray-100 border-b border-gray-300 px-3 flex items-center justify-between flex-shrink-0 group">
            {renamingWorkbook === workbookName ? (
              <div className="flex items-center gap-2 flex-1">
                <input
                  type="text"
                  value={newWorkbookName}
                  onChange={(e) => setNewWorkbookName(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      renameWorkbook(workbookName, newWorkbookName);
                    } else if (e.key === "Escape") {
                      cancelWorkbookRenaming();
                    }
                  }}
                  onBlur={() => {
                    if (
                      newWorkbookName.trim() &&
                      newWorkbookName.trim() !== workbookName
                    ) {
                      renameWorkbook(workbookName, newWorkbookName);
                    } else {
                      cancelWorkbookRenaming();
                    }
                  }}
                  className="text-sm font-medium bg-white border border-blue-500 rounded px-2 py-1 outline-none flex-1"
                  autoFocus
                  data-testid={`rename-workbook-input-${workbookName}`}
                />
              </div>
            ) : (
              <div className="flex items-center gap-2 flex-1">
                <span
                  className="text-sm font-medium text-gray-700 cursor-pointer hover:text-gray-900"
                  onDoubleClick={() => startRenamingWorkbook(workbookName)}
                  data-testid={`workbook-name-${workbookName}`}
                >
                  {workbookName}
                </span>
                <div className="opacity-0 group-hover:opacity-100 transition-opacity flex items-center gap-1">
                  <div
                    className="h-4 w-4 p-0 text-gray-500 hover:text-blue-600 flex items-center justify-center cursor-pointer"
                    onClick={() => startRenamingWorkbook(workbookName)}
                    title="Rename Workbook"
                    data-testid={`rename-workbook-${workbookName}`}
                  >
                    <Edit2 className="h-3 w-3" />
                  </div>
                  <div
                    className="h-4 w-4 p-0 text-gray-500 hover:text-green-600 flex items-center justify-center cursor-pointer"
                    onClick={() => cloneWorkbook(workbookName)}
                    title="Clone Workbook"
                    data-testid={`clone-workbook-${workbookName}`}
                  >
                    <Copy className="h-3 w-3" />
                  </div>
                  {workbookGridItems.length > 1 && (
                    <div
                      className="h-4 w-4 p-0 text-gray-500 hover:text-red-600 flex items-center justify-center cursor-pointer"
                      onClick={() => deleteWorkbook(workbookName)}
                      title="Delete Workbook"
                      data-testid={`delete-workbook-${workbookName}`}
                    >
                      <Trash2 className="h-3 w-3" />
                    </div>
                  )}
                </div>
              </div>
            )}
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
                        onClick={(ev) => {
                          return updateWorkbookActiveSheet(
                            workbookName,
                            sheetName
                          );
                        }}
                        onDoubleClick={(ev) => {
                          return startRenaming(workbookName, sheetName);
                        }}
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
      renamingWorkbook,
      newWorkbookName,
      renameWorkbook,
      startRenamingWorkbook,
      cancelWorkbookRenaming,
      cloneWorkbook,
      deleteWorkbook,
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

  // Show loading spinner while initializing
  if (isLoading) {
    return (
      <div className="h-full flex items-center justify-center bg-gray-50">
        <div className="text-center">
          <div className="inline-block animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
          <p className="text-sm text-gray-600">Loading workbook from OPFS...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col">
      {/* Grid-based header */}
      <div className="border-b border-gray-200 bg-gray-50">
        <div className="p-3 border-b border-gray-200">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <h1 className="text-lg font-semibold text-gray-800">
                FormulaEngine Multi-Workbook Demo
              </h1>
              {currentFileName && (
                <div className="flex items-center gap-2 px-3 py-1 bg-blue-50 rounded border border-blue-200">
                  <FileText className="h-4 w-4 text-blue-600" />
                  <span className="text-sm font-medium text-blue-800" data-testid="current-filename">
                    {currentFileName}
                  </span>
                  {hasUnsavedChanges && (
                    <span className="text-xs text-orange-600">●</span>
                  )}
                </div>
              )}
            </div>
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
                {/* File Operations */}
                <div className="flex items-center gap-1 border-r border-gray-300 pr-2">
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={() => setShowFileManager(!showFileManager)}
                    data-testid="file-manager-button"
                    title="File Manager"
                  >
                    <Files className="h-4 w-4 mr-1" />
                    Files ({opfsFiles.length})
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={newFile}
                    data-testid="new-file-button"
                    title="New File in OPFS"
                  >
                    <File className="h-4 w-4 mr-1" />
                    New
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
                    onClick={saveFile}
                    data-testid="save-button"
                    title="Save to OPFS (auto-saves after 2s)"
                  >
                    <Save className="h-4 w-4 mr-1" />
                    {hasUnsavedChanges ? "Save" : "Saved"}
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={renameFile}
                    data-testid="rename-file-button"
                    title="Rename current file"
                  >
                    <Edit2 className="h-4 w-4 mr-1" />
                    Rename
                  </Button>
                </div>

                {/* Import/Export to/from File System */}
                <div className="flex items-center gap-1 border-r border-gray-300 pr-2">
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={exportToFileSystem}
                    data-testid="export-button"
                    title="Export current file to your file system"
                  >
                    <Download className="h-4 w-4 mr-1" />
                    Export
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={importFromFileSystem}
                    data-testid="import-button"
                    title="Import file from your file system to OPFS"
                  >
                    <Upload className="h-4 w-4 mr-1" />
                    Import
                  </Button>
                </div>

                {/* Workbook Operations */}
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
                
                {/* Tools */}
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
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* File Manager Panel */}
      {showFileManager && (
        <div
          className="border-b border-gray-200 bg-gray-50 p-4"
          data-testid="file-manager-panel"
        >
          <div className="bg-white p-3 rounded border border-gray-200">
            <h3 className="text-sm font-semibold text-gray-800 mb-3">
              OPFS Files ({opfsFiles.length})
            </h3>
            <div className="space-y-2 max-h-60 overflow-y-auto">
              {opfsFiles.length === 0 ? (
                <p className="text-xs text-gray-500 italic">
                  No files in OPFS. Create a new file to get started.
                </p>
              ) : (
                opfsFiles.map((filename) => (
                  <div
                    key={filename}
                    className={`
                      flex items-center justify-between p-2 rounded cursor-pointer transition-colors
                      ${
                        filename === currentFileName
                          ? "bg-blue-100 border border-blue-300"
                          : "bg-gray-50 hover:bg-gray-100 border border-gray-200"
                      }
                    `}
                    data-testid={`opfs-file-${filename}`}
                  >
                    <div
                      className="flex items-center gap-2 flex-1"
                      onClick={() => {
                        if (filename !== currentFileName) {
                          if (
                            hasUnsavedChanges &&
                            !confirm(
                              "You have unsaved changes. Switch files anyway?"
                            )
                          ) {
                            return;
                          }
                          loadFile(filename);
                        }
                      }}
                    >
                      <FileText
                        className={`h-4 w-4 ${
                          filename === currentFileName
                            ? "text-blue-600"
                            : "text-gray-600"
                        }`}
                      />
                      <span
                        className={`text-sm ${
                          filename === currentFileName
                            ? "font-semibold text-blue-800"
                            : "text-gray-700"
                        }`}
                      >
                        {filename}
                      </span>
                      {filename === currentFileName && (
                        <span className="text-xs px-2 py-0.5 bg-blue-500 text-white rounded">
                          Current
                        </span>
                      )}
                    </div>
                    {filename !== currentFileName && opfsFiles.length > 1 && (
                      <Button
                        size="sm"
                        variant="ghost"
                        className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                        onClick={(e) => {
                          e.stopPropagation();
                          deleteFile(filename);
                        }}
                        title="Delete File"
                        data-testid={`delete-opfs-file-${filename}`}
                      >
                        <Trash2 className="h-3 w-3" />
                      </Button>
                    )}
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}

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
                              <div className="flex items-center gap-1">
                                <Button
                                  size="sm"
                                  variant="ghost"
                                  className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                                  onClick={() => {
                                    if (
                                      confirm(
                                        `Are you sure you want to delete table "${tableName}"?`
                                      )
                                    ) {
                                      deleteTable(tableName, workbookName);
                                    }
                                  }}
                                  title="Delete Table"
                                  data-testid={`delete-table-${tableName}`}
                                >
                                  <Trash2 className="h-3 w-3" />
                                </Button>
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
