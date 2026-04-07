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
  History,
  Plus,
  Redo,
  Save,
  Trash2,
  Undo,
  Upload,
  X,
} from "lucide-react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { FormulaEngine } from "../src/core/engine";
import { deserialize, serialize } from "../src/core/map-serializer";
import { hexToLch, lchToHex } from "../src/core/utils/color-utils";
import { useEngine } from "../src/react/hooks";
import type { EngineAction } from "../src/core/commands/types";
import { SpreadsheetWithFormulaBar } from "./components/SpreadsheetWithFormulaBar";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { WorkbookClipboardManager } from "./WorkbookClipboardManager";

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
  engineStateSerialized: string; // Serialized engine state as string
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
  const clipboardManager = new WorkbookClipboardManager(engine);

  // Create first workbook and sheet with sample data
  const workbookName = "Workbook1";
  engine.addWorkbook(workbookName);
  const sheetNameToAdd = "Sheet1";
  engine.addSheet({
    workbookName,
    sheetName: sheetNameToAdd,
  });
  const sheetName = sheetNameToAdd;

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
    clipboardManager,
  };
};

const {
  engine,
  workbookGridItems: initialWorkbookGridItems,
  viewport: initialViewport,
  clipboardManager,
} = createEngine();
console.log("engine", engine);

// Helper function to format action types for display
function formatAction(action: EngineAction): {
  type: string;
  description?: string;
  details?: string;
} {
  const payload = action.payload as any;
  
  switch (action.type) {
    case "SET_CELL_CONTENT":
      return {
        type: "Set Cell",
        description: `Cell ${payload.address?.sheetName || ""} ${getCellRef(payload.address)}`,
        details: String(payload.content || "").substring(0, 50),
      };
    case "SET_SHEET_CONTENT":
      return {
        type: "Set Sheet Content",
        description: `${payload.opts?.sheetName || ""} (${payload.content?.length || 0} cells)`,
      };
    case "CLEAR_RANGE":
      return {
        type: "Clear Range",
        description: `${payload.address?.sheetName || ""} ${formatRange(payload.address?.range)}`,
      };
    case "AUTO_FILL":
      return {
        type: "Auto Fill",
        description: `${payload.opts?.sheetName || ""}`,
      };
    case "PASTE_CELLS":
      return {
        type: "Paste Cells",
        description: `${payload.target?.sheetName || ""}`,
      };
    case "FILL_AREAS":
      return {
        type: "Fill Areas",
        description: `${payload.targetRanges?.length || 0} ranges`,
      };
    case "MOVE_CELL":
      return {
        type: "Move Cell",
        description: `${payload.source?.sheetName || ""} → ${payload.target?.sheetName || ""}`,
      };
    case "MOVE_RANGE":
      return {
        type: "Move Range",
        description: `${payload.sourceRange?.sheetName || ""} → ${payload.target?.sheetName || ""}`,
      };
    case "ADD_WORKBOOK":
      return {
        type: "Add Workbook",
        description: payload.workbookName || "",
      };
    case "REMOVE_WORKBOOK":
      return {
        type: "Remove Workbook",
        description: payload.workbookName || "",
      };
    case "RENAME_WORKBOOK":
      return {
        type: "Rename Workbook",
        description: `${payload.workbookName || ""} → ${payload.newWorkbookName || ""}`,
      };
    case "CLONE_WORKBOOK":
      return {
        type: "Clone Workbook",
        description: `${payload.fromWorkbookName || ""} → ${payload.toWorkbookName || ""}`,
      };
    case "ADD_SHEET":
      return {
        type: "Add Sheet",
        description: `${payload.workbookName || ""} → ${payload.sheetName || ""}`,
      };
    case "REMOVE_SHEET":
      return {
        type: "Remove Sheet",
        description: `${payload.workbookName || ""} → ${payload.sheetName || ""}`,
      };
    case "RENAME_SHEET":
      return {
        type: "Rename Sheet",
        description: `${payload.workbookName || ""} → ${payload.sheetName || ""} → ${payload.newSheetName || ""}`,
      };
    case "ADD_TABLE":
      return {
        type: "Add Table",
        description: `${payload.tableName || ""} (${payload.workbookName || ""})`,
      };
    case "REMOVE_TABLE":
      return {
        type: "Remove Table",
        description: `${payload.tableName || ""} (${payload.workbookName || ""})`,
      };
    case "RENAME_TABLE":
      return {
        type: "Rename Table",
        description: `${payload.oldName || ""} → ${payload.newName || ""}`,
      };
    case "UPDATE_TABLE":
      return {
        type: "Update Table",
        description: payload.tableName || "",
      };
    case "RESET_TABLES":
      return {
        type: "Reset Tables",
        description: `${payload.tables?.length || 0} tables`,
      };
    case "ADD_NAMED_EXPRESSION":
      return {
        type: "Add Named Expression",
        description: payload.expressionName || "",
      };
    case "REMOVE_NAMED_EXPRESSION":
      return {
        type: "Remove Named Expression",
        description: payload.expressionName || "",
      };
    case "UPDATE_NAMED_EXPRESSION":
      return {
        type: "Update Named Expression",
        description: payload.expressionName || "",
      };
    case "RENAME_NAMED_EXPRESSION":
      return {
        type: "Rename Named Expression",
        description: `${payload.oldName || ""} → ${payload.newName || ""}`,
      };
    case "SET_NAMED_EXPRESSIONS":
      return {
        type: "Set Named Expressions",
        description: `${payload.expressions?.length || 0} expressions`,
      };
    case "ADD_CONDITIONAL_STYLE":
      return {
        type: "Add Conditional Style",
        description: `${payload.style?.areas?.[0]?.workbookName || ""}`,
      };
    case "REMOVE_CONDITIONAL_STYLE":
      return {
        type: "Remove Conditional Style",
        description: payload.workbookName || "",
      };
    case "ADD_CELL_STYLE":
      return {
        type: "Add Cell Style",
        description: `${payload.style?.areas?.[0]?.workbookName || ""}`,
      };
    case "REMOVE_CELL_STYLE":
      return {
        type: "Remove Cell Style",
        description: payload.workbookName || "",
      };
    case "CLEAR_CELL_STYLES":
      return {
        type: "Clear Cell Styles",
        description: formatRange(payload.range),
      };
    case "SET_CELL_METADATA":
      return {
        type: "Set Cell Metadata",
        description: getCellRef(payload.address),
      };
    case "SET_SHEET_METADATA":
      return {
        type: "Set Sheet Metadata",
        description: `${payload.opts?.workbookName || ""} → ${payload.opts?.sheetName || ""}`,
      };
    case "SET_WORKBOOK_METADATA":
      return {
        type: "Set Workbook Metadata",
        description: payload.workbookName || "",
      };
    case "RESET_TO_SERIALIZED":
      return {
        type: "Reset to Serialized State",
        description: "Full state reset",
      };
    default:
      return {
        type: action.type.replace(/_/g, " "),
        description: JSON.stringify(payload).substring(0, 100),
      };
  }
}

function getCellRef(address: any): string {
  if (!address) return "";
  const col = String.fromCharCode(65 + (address.colIndex || 0));
  const row = (address.rowIndex || 0) + 1;
  return `${col}${row}`;
}

function formatRange(range: any): string {
  if (!range) return "";
  const startCol = String.fromCharCode(65 + (range.start?.col || 0));
  const startRow = (range.start?.row || 0) + 1;
  const endCol = range.end?.col?.type === "number"
    ? String.fromCharCode(65 + range.end.col.value)
    : "∞";
  const endRow = range.end?.row?.type === "number"
    ? range.end.row.value + 1
    : "∞";
  return `${startCol}${startRow}:${endCol}${endRow}`;
}

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
  const [showChangelog, setShowChangelog] = useState(false);
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

  // Conditional Formatting UI state
  const [showConditionalFormatting, setShowConditionalFormatting] =
    useState(false);
  const [newStyleArea, setNewStyleArea] = useState("");
  const [newStyleType, setNewStyleType] = useState<"formula" | "gradient">(
    "formula"
  );
  const [newStyleFormula, setNewStyleFormula] = useState("ROW() > 4");
  const [newStyleColorL, setNewStyleColorL] = useState("70");
  const [newStyleColorC, setNewStyleColorC] = useState("80");
  const [newStyleColorH, setNewStyleColorH] = useState("0");
  const [newStyleGradientType, setNewStyleGradientType] = useState<
    "lowest_highest" | "number"
  >("lowest_highest");
  const [newStyleMinColorL, setNewStyleMinColorL] = useState("90");
  const [newStyleMinColorC, setNewStyleMinColorC] = useState("10");
  const [newStyleMinColorH, setNewStyleMinColorH] = useState("120");
  const [newStyleMaxColorL, setNewStyleMaxColorL] = useState("30");
  const [newStyleMaxColorC, setNewStyleMaxColorC] = useState("80");
  const [newStyleMaxColorH, setNewStyleMaxColorH] = useState("0");
  const [newStyleMinFormula, setNewStyleMinFormula] = useState("0");
  const [newStyleMaxFormula, setNewStyleMaxFormula] = useState("100");
  const [editingStyle, setEditingStyle] = useState<{
    workbookName: string;
    index: number;
  } | null>(null);

  // Cell Styles UI state
  const [newCellStyleArea, setNewCellStyleArea] = useState("");
  const [newCellStyleBackgroundColor, setNewCellStyleBackgroundColor] =
    useState("#FFFFFF");
  const [newCellStyleColor, setNewCellStyleColor] = useState("#000000");
  const [newCellStyleFontSize, setNewCellStyleFontSize] = useState("12");
  const [newCellStyleBold, setNewCellStyleBold] = useState(false);
  const [newCellStyleItalic, setNewCellStyleItalic] = useState(false);
  const [newCellStyleUnderline, setNewCellStyleUnderline] = useState(false);
  const [editingCellStyle, setEditingCellStyle] = useState<{
    workbookName: string;
    index: number;
  } | null>(null);

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
  const [actionLog, setActionLog] = useState(engine.getActionLog());

  const applyLoadedState = useCallback(
    (filename: string, data: SavedState) => {
      if (!data.engineStateSerialized || !data.workbookGridItems) {
        throw new Error("Invalid file format");
      }

      engine.resetToSerializedEngine(data.engineStateSerialized);
      setWorkbookGridItems(data.workbookGridItems);
      _setViewport(data.viewport);
      setCurrentFileName(filename);
      setHasUnsavedChanges(false);
      console.log(`Loaded ${filename}`);
    },
    [engine, _setViewport]
  );

  // Load from OPFS on mount
  useEffect(() => {
    const loadInitialFile = async () => {
      try {
        const files = await listOPFSFiles();
        setOpfsFiles(files);

        // Try to load the default file or the first available file
        const fileToLoad = files.includes(DEFAULT_FILE)
          ? DEFAULT_FILE
          : files[0];

        if (fileToLoad) {
          const data = await loadFromOPFS(fileToLoad);
          if (data) {
            applyLoadedState(fileToLoad, data);
          }
        }
      } catch (error) {
        console.error("Failed to load from OPFS:", error);
      } finally {
        setIsLoading(false);
      }
    };

    loadInitialFile();
  }, [applyLoadedState]);

  // Save to OPFS
  const saveFile = useCallback(async () => {
    try {
      const dataToSave: SavedState = {
        workbookGridItems,
        engineStateSerialized: engine.serializeEngine(),
        viewport,
      };

      await saveToOPFS(currentFileName, dataToSave);
      setHasUnsavedChanges(false);
      console.log(`Saved to OPFS: ${currentFileName}`);

      // Refresh file list
      const files = await listOPFSFiles();
      setOpfsFiles(files);
    } catch (error) {
      console.error("Failed to save to OPFS:", error);
      alert(
        `Failed to save: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }, [workbookGridItems, engine, viewport, currentFileName]);

  // Load a file from OPFS
  const loadFile = useCallback(
    async (filename: string) => {
      try {
        const data = await loadFromOPFS(filename);
        if (!data) {
          throw new Error("Invalid file format");
        }
        applyLoadedState(filename, data);
      } catch (error) {
        console.error(`Failed to load ${filename}:`, error);
        alert(
          `Failed to load file: ${
            error instanceof Error ? error.message : "Unknown error"
          }`
        );
      }
    },
    [applyLoadedState]
  );

  // Create new file
  const newFile = useCallback(async () => {
    if (hasUnsavedChanges) {
      if (
        !confirm(
          "You have unsaved changes. Are you sure you want to create a new file?"
        )
      ) {
        return;
      }
    }

    // Prompt for new filename
    const newFileName = prompt(
      "Enter new file name (without .json extension):",
      "workbook"
    );
    if (!newFileName) return;

    const fullFileName = newFileName.endsWith(".json")
      ? newFileName
      : `${newFileName}.json`;

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
    const sheetNameToAdd = "Sheet1";
    engine.addSheet({
      workbookName,
      sheetName: sheetNameToAdd,
    });
    const sheetName = sheetNameToAdd;

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

    engine._evaluationManager.clearEvaluationCache();
    engine._eventManager.emitUpdate();

    // Save the new file immediately
    const dataToSave: SavedState = {
      workbookGridItems: [defaultWorkbookItem],
      engineStateSerialized: engine.serializeEngine(),
      viewport: undefined,
    };
    await saveToOPFS(fullFileName, dataToSave);
    setHasUnsavedChanges(false);

    // Refresh file list
    const files = await listOPFSFiles();
    setOpfsFiles(files);
  }, [engine, hasUnsavedChanges, _setViewport, opfsFiles]);

  // Delete a file from OPFS
  const deleteFile = useCallback(
    async (filename: string) => {
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
        alert(
          `Failed to delete file: ${
            error instanceof Error ? error.message : "Unknown error"
          }`
        );
      }
    },
    [currentFileName, newFile]
  );

  // Rename current file
  const renameFile = useCallback(async () => {
    const newName = prompt(
      "Enter new file name (without .json extension):",
      currentFileName.replace(".json", "")
    );
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
      alert(
        `Failed to rename file: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }, [currentFileName, opfsFiles]);

  // Export current file to file system
  const exportToFileSystem = useCallback(() => {
    try {
      const dataToExport: SavedState = {
        workbookGridItems,
        engineStateSerialized: engine.serializeEngine(),
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
          if (
            !importedData.engineStateSerialized ||
            !importedData.workbookGridItems
          ) {
            throw new Error("Invalid file format: missing required fields");
          }

          // Save to OPFS with the imported filename
          let filename = file.name;
          if (opfsFiles.includes(filename)) {
            if (
              !confirm(`File "${filename}" already exists in OPFS. Overwrite?`)
            ) {
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
          applyLoadedState(filename, importedData);

          // Refresh file list
          const files = await listOPFSFiles();
          setOpfsFiles(files);

          console.log(`Imported ${file.name} to OPFS as ${filename}`);
          alert(`File imported successfully as "${filename}"`);
        } catch (error) {
          console.error("Failed to import file:", error);
          alert(
            `Failed to import file: ${
              error instanceof Error ? error.message : "Unknown error"
            }`
          );
        }
      };

      reader.readAsText(file);
    };

    input.click();
  }, [applyLoadedState, opfsFiles]);

  // Mark as having unsaved changes when sheets change
  const markUnsavedChanges = useCallback(() => {
    setHasUnsavedChanges(true);
  }, []);

  // Track changes
  useEffect(() => {
    const unsubscribe = engine.onUpdate(() => {
      markUnsavedChanges();
      setActionLog(engine.getActionLog());
    });
    return unsubscribe;
  }, [engine, markUnsavedChanges]);

  // Keyboard shortcuts for undo/redo
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) {
        e.preventDefault();
        if (engine.canUndo()) {
          engine.undo();
          setActionLog(engine.getActionLog());
        }
      } else if (
        (e.ctrlKey || e.metaKey) &&
        (e.key === "y" || (e.key === "z" && e.shiftKey))
      ) {
        e.preventDefault();
        if (engine.canRedo()) {
          engine.redo();
          setActionLog(engine.getActionLog());
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [engine]);

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
        const workbookItem = prev.find((item) => item.name === workbookName);

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
      const workbookItem = workbookGridItems.find(
        (item) => item.name === workbookName
      );
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

        engine.removeNamedExpression({
          expressionName: name,
          sheetName: isGlobal ? undefined : activeSheet,
          workbookName: isGlobal ? undefined : targetWorkbook,
        });
        markUnsavedChanges();
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
        const sheetNames = engine.getOrderedSheetNames(workbookName);
        if (sheetNames.length <= 1) {
          alert("Cannot delete the last sheet in a workbook");
          return;
        }

        engine.removeSheet({ workbookName, sheetName });

        // If we deleted the active sheet, switch to the first available sheet
        const workbookItem = workbookGridItems.find(
          (item) => item.name === workbookName
        );
        if (workbookItem?.activeSheet === sheetName) {
          const firstSheet = engine.getOrderedSheetNames(workbookName)[0];
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
          `Failed to clone workbook: ${
            error instanceof Error ? error.message : "Unknown error"
          }`
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
          `Failed to delete workbook: ${
            error instanceof Error ? error.message : "Unknown error"
          }`
        );
      }
    },
    [workbookGridItems, engine, markUnsavedChanges]
  );

  // Delete table
  const deleteTable = useCallback(
    (tableName: string, workbookName: string) => {
      try {
        engine.removeTable({ tableName, workbookName });
        markUnsavedChanges();
      } catch (error) {
        console.error("Failed to delete table:", error);
      }
    },
    [engine, markUnsavedChanges]
  );

  // Reset style form to default values
  const resetStyleForm = useCallback(() => {
    setNewStyleArea("");
    setNewStyleType("formula");
    setNewStyleFormula("ROW() > 4");
    setNewStyleColorL("70");
    setNewStyleColorC("80");
    setNewStyleColorH("0");
    setNewStyleGradientType("lowest_highest");
    setNewStyleMinColorL("90");
    setNewStyleMinColorC("10");
    setNewStyleMinColorH("120");
    setNewStyleMaxColorL("30");
    setNewStyleMaxColorC("80");
    setNewStyleMaxColorH("0");
    setNewStyleMinFormula("0");
    setNewStyleMaxFormula("100");
    setEditingStyle(null);
  }, []);

  // Add conditional style
  const addConditionalStyle = useCallback(() => {
    try {
      // Parse the range format: [workbook]'sheet'!A1:C10 or just A1:C10
      let workbookName: string;
      let sheetName: string;
      let rangeStr: string;

      // Check if the input includes workbook and sheet references
      const fullRangeMatch = newStyleArea.match(
        /^\[([^\]]+)\](?:'([^']+(?:''[^']*)*)'|([^!]+))!(.+)$/
      );

      if (fullRangeMatch) {
        // Format: [workbook]'sheet'!range or [workbook]sheet!range
        workbookName = fullRangeMatch[1]!;
        sheetName = fullRangeMatch[2]
          ? fullRangeMatch[2].replace(/''/g, "'")
          : fullRangeMatch[3]!;
        rangeStr = fullRangeMatch[4]!;
      } else {
        // Simple format: just the range, use default workbook/sheet
        workbookName = workbookGridItems[0]?.name!;
        sheetName = workbookGridItems[0]?.activeSheet!;
        rangeStr = newStyleArea;

        if (!workbookName || !sheetName) {
          alert("No active workbook or sheet");
          return;
        }
      }

      const colToIndex = (col: string): number => {
        let result = 0;
        for (let i = 0; i < col.length; i++) {
          result = result * 26 + (col.charCodeAt(i) - 64);
        }
        return result - 1;
      };

      // Parse the range part - supports multiple formats
      let startCol: number, startRow: number;
      let endCol: number, endRow: number;

      // Match different range formats
      // A5:D10 (closed), A5:10 (row-bounded), A5:D (col-bounded), A5:INFINITY (open both)
      const closedMatch = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      const rowBoundedMatch = rangeStr.match(/^([A-Z]+)(\d+):(\d+)$/);
      const colBoundedMatch = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)$/);
      const openBothMatch = rangeStr.match(/^([A-Z]+)(\d+):INFINITY$/);

      if (closedMatch) {
        // Closed rectangle: A5:D10
        startCol = colToIndex(closedMatch[1]!);
        startRow = parseInt(closedMatch[2]!) - 1;
        endCol = colToIndex(closedMatch[3]!);
        endRow = parseInt(closedMatch[4]!) - 1;
      } else if (rowBoundedMatch) {
        // Row-bounded (col-infinite): A5:10
        startCol = colToIndex(rowBoundedMatch[1]!);
        startRow = parseInt(rowBoundedMatch[2]!) - 1;
        endCol = Infinity; // Will be set as infinity type below
        endRow = parseInt(rowBoundedMatch[3]!) - 1;
      } else if (colBoundedMatch) {
        // Col-bounded (row-infinite): A5:D
        startCol = colToIndex(colBoundedMatch[1]!);
        startRow = parseInt(colBoundedMatch[2]!) - 1;
        endCol = colToIndex(colBoundedMatch[3]!);
        endRow = Infinity; // Will be set as infinity type below
      } else if (openBothMatch) {
        // Open both: A5:INFINITY
        startCol = colToIndex(openBothMatch[1]!);
        startRow = parseInt(openBothMatch[2]!) - 1;
        endCol = Infinity;
        endRow = Infinity;
      } else {
        alert(
          "Invalid range format. Use format like [Workbook]Sheet!A1:C10 or A1:C10"
        );
        return;
      }

      // Normalize the range: ensure start is always before end
      // Swap columns if startCol > endCol (only if both are finite numbers)
      if (
        typeof startCol === "number" &&
        typeof endCol === "number" &&
        isFinite(startCol) &&
        isFinite(endCol) &&
        startCol > endCol
      ) {
        [startCol, endCol] = [endCol, startCol];
      }
      // Swap rows if startRow > endRow (only if both are finite numbers)
      if (
        typeof startRow === "number" &&
        typeof endRow === "number" &&
        isFinite(startRow) &&
        isFinite(endRow) &&
        startRow > endRow
      ) {
        [startRow, endRow] = [endRow, startRow];
      }

      let condition;
      if (newStyleType === "formula") {
        condition = {
          type: "formula" as const,
          formula: newStyleFormula,
          color: {
            l: parseFloat(newStyleColorL),
            c: parseFloat(newStyleColorC),
            h: parseFloat(newStyleColorH),
          },
        };
      } else {
        if (newStyleGradientType === "lowest_highest") {
          condition = {
            type: "gradient" as const,
            min: {
              type: "lowest_value" as const,
              color: {
                l: parseFloat(newStyleMinColorL),
                c: parseFloat(newStyleMinColorC),
                h: parseFloat(newStyleMinColorH),
              },
            },
            max: {
              type: "highest_value" as const,
              color: {
                l: parseFloat(newStyleMaxColorL),
                c: parseFloat(newStyleMaxColorC),
                h: parseFloat(newStyleMaxColorH),
              },
            },
          };
        } else {
          condition = {
            type: "gradient" as const,
            min: {
              type: "number" as const,
              color: {
                l: parseFloat(newStyleMinColorL),
                c: parseFloat(newStyleMinColorC),
                h: parseFloat(newStyleMinColorH),
              },
              valueFormula: newStyleMinFormula,
            },
            max: {
              type: "number" as const,
              color: {
                l: parseFloat(newStyleMaxColorL),
                c: parseFloat(newStyleMaxColorC),
                h: parseFloat(newStyleMaxColorH),
              },
              valueFormula: newStyleMaxFormula,
            },
          };
        }
      }

      // If editing, remove the old style first, then add the new one
      if (editingStyle && editingStyle.workbookName === workbookName) {
        // Remove the old style
        engine.removeConditionalStyle(
          editingStyle.workbookName,
          editingStyle.index
        );
        // If the index would change after removal, adjust it
        const insertIndex = editingStyle.index;
        // Add the new style (it will be added at the end, so we need to handle ordering)
        // For simplicity, we'll just add it and the order might change slightly
        engine.addConditionalStyle({
          areas: [{
            workbookName,
            sheetName,
            range: {
              start: { col: startCol, row: startRow },
              end: {
                col:
                  endCol === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endCol },
                row:
                  endRow === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endRow },
              },
            },
          }],
          condition,
        });
        setEditingStyle(null);
        resetStyleForm();
      } else {
        engine.addConditionalStyle({
          areas: [{
            workbookName,
            sheetName,
            range: {
              start: { col: startCol, row: startRow },
              end: {
                col:
                  endCol === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endCol },
                row:
                  endRow === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endRow },
              },
            },
          }],
          condition,
        });
      }

      markUnsavedChanges();
    } catch (error) {
      console.error("Failed to add conditional style:", error);
      alert(
        `Failed to add conditional style: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }, [
    workbookGridItems,
    engine,
    markUnsavedChanges,
    newStyleArea,
    newStyleType,
    newStyleFormula,
    newStyleColorL,
    newStyleColorC,
    newStyleColorH,
    newStyleGradientType,
    newStyleMinColorL,
    newStyleMinColorC,
    newStyleMinColorH,
    newStyleMaxColorL,
    newStyleMaxColorC,
    newStyleMaxColorH,
    newStyleMinFormula,
    newStyleMaxFormula,
    editingStyle,
    resetStyleForm,
  ]);

  // Delete conditional style
  const deleteConditionalStyle = useCallback(
    (workbookName: string, index: number) => {
      try {
        engine.removeConditionalStyle(workbookName, index);
        markUnsavedChanges();
        // If we're editing this style, cancel editing
        if (
          editingStyle &&
          editingStyle.workbookName === workbookName &&
          editingStyle.index === index
        ) {
          setEditingStyle(null);
          resetStyleForm();
        }
      } catch (error) {
        console.error("Failed to delete conditional style:", error);
      }
    },
    [engine, markUnsavedChanges, editingStyle, resetStyleForm]
  );

  // Load style into form for editing
  const editConditionalStyle = useCallback(
    (workbookName: string, index: number) => {
      // Get style directly from engineState since index is global
      const style = engineState.conditionalStyles?.[index];
      if (!style) return;

      // Convert range to canonical format
      const colToLetter = (col: number): string => {
        let result = "";
        let c = col;
        while (c >= 0) {
          result = String.fromCharCode(65 + (c % 26)) + result;
          c = Math.floor(c / 26) - 1;
        }
        return result;
      };

      const startCol = colToLetter(style.areas[0]!.range.start.col);
      const startRow = style.areas[0]!.range.start.row + 1;
      const isColInfinity = style.areas[0]!.range.end.col.type === "infinity";
      const isRowInfinity = style.areas[0]!.range.end.row.type === "infinity";

      let cellRange: string;
      if (isColInfinity && isRowInfinity) {
        cellRange = `${startCol}${startRow}:INFINITY`;
      } else if (isColInfinity) {
        const endRow =
          style.areas[0]!.range.end.row.type === "number"
            ? style.areas[0]!.range.end.row.value + 1
            : startRow;
        cellRange = `${startCol}${startRow}:${endRow}`;
      } else if (isRowInfinity) {
        const endCol =
          style.areas[0]!.range.end.col.type === "number"
            ? colToLetter(style.areas[0]!.range.end.col.value)
            : startCol;
        cellRange = `${startCol}${startRow}:${endCol}`;
      } else {
        const endCol =
          style.areas[0]!.range.end.col.type === "number"
            ? colToLetter(style.areas[0]!.range.end.col.value)
            : startCol;
        const endRow =
          style.areas[0]!.range.end.row.type === "number"
            ? style.areas[0]!.range.end.row.value + 1
            : startRow;
        cellRange = `${startCol}${startRow}:${endCol}${endRow}`;
      }

      // Quote sheet name if needed
      const needsQuotes = /[ '!]/.test(style.areas[0]!.sheetName);
      const sheetRef = needsQuotes
        ? `'${style.areas[0]!.sheetName.replace(/'/g, "''")}'`
        : style.areas[0]!.sheetName;
      const rangeStr = `[${style.areas[0]!.workbookName}]${sheetRef}!${cellRange}`;

      setNewStyleArea(rangeStr);

      if (style.condition.type === "formula") {
        setNewStyleType("formula");
        setNewStyleFormula(style.condition.formula);
        setNewStyleColorL(style.condition.color.l.toFixed(0));
        setNewStyleColorC(style.condition.color.c.toFixed(0));
        setNewStyleColorH(style.condition.color.h.toFixed(0));
      } else {
        setNewStyleType("gradient");
        setNewStyleMinColorL(style.condition.min.color.l.toFixed(0));
        setNewStyleMinColorC(style.condition.min.color.c.toFixed(0));
        setNewStyleMinColorH(style.condition.min.color.h.toFixed(0));
        setNewStyleMaxColorL(style.condition.max.color.l.toFixed(0));
        setNewStyleMaxColorC(style.condition.max.color.c.toFixed(0));
        setNewStyleMaxColorH(style.condition.max.color.h.toFixed(0));

        if (
          style.condition.min.type === "lowest_value" &&
          style.condition.max.type === "highest_value"
        ) {
          setNewStyleGradientType("lowest_highest");
        } else {
          setNewStyleGradientType("number");
          if (style.condition.min.type === "number") {
            setNewStyleMinFormula(style.condition.min.valueFormula);
          }
          if (style.condition.max.type === "number") {
            setNewStyleMaxFormula(style.condition.max.valueFormula);
          }
        }
      }

      setEditingStyle({ workbookName, index });
    },
    [engineState]
  );

  // Reset cell style form
  const resetCellStyleForm = useCallback(() => {
    setNewCellStyleArea("");
    setNewCellStyleBackgroundColor("#FFFFFF");
    setNewCellStyleColor("#000000");
    setNewCellStyleFontSize("12");
    setNewCellStyleBold(false);
    setNewCellStyleItalic(false);
    setNewCellStyleUnderline(false);
    setEditingCellStyle(null);
  }, []);

  // Add cell style
  const addCellStyle = useCallback(() => {
    try {
      // Parse the range format: [workbook]'sheet'!A1:C10 or just A1:C10
      let workbookName: string;
      let sheetName: string;
      let rangeStr: string;

      // Check if the input includes workbook and sheet references
      const fullRangeMatch = newCellStyleArea.match(
        /^\[([^\]]+)\](?:'([^']+(?:''[^']*)*)'|([^!]+))!(.+)$/
      );

      if (fullRangeMatch) {
        // Format: [workbook]'sheet'!range or [workbook]sheet!range
        workbookName = fullRangeMatch[1]!;
        sheetName = fullRangeMatch[2]
          ? fullRangeMatch[2].replace(/''/g, "'")
          : fullRangeMatch[3]!;
        rangeStr = fullRangeMatch[4]!;
      } else {
        // Simple format: just the range, use default workbook/sheet
        workbookName = workbookGridItems[0]?.name!;
        sheetName = workbookGridItems[0]?.activeSheet!;
        rangeStr = newCellStyleArea;

        if (!workbookName || !sheetName) {
          alert("No active workbook or sheet");
          return;
        }
      }

      const colToIndex = (col: string): number => {
        let result = 0;
        for (let i = 0; i < col.length; i++) {
          result = result * 26 + (col.charCodeAt(i) - 64);
        }
        return result - 1;
      };

      // Parse the range part - supports multiple formats
      let startCol: number, startRow: number;
      let endCol: number, endRow: number;

      // Match different range formats
      const closedMatch = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      const rowBoundedMatch = rangeStr.match(/^([A-Z]+)(\d+):(\d+)$/);
      const colBoundedMatch = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)$/);
      const openBothMatch = rangeStr.match(/^([A-Z]+)(\d+):INFINITY$/);

      if (closedMatch) {
        startCol = colToIndex(closedMatch[1]!);
        startRow = parseInt(closedMatch[2]!) - 1;
        endCol = colToIndex(closedMatch[3]!);
        endRow = parseInt(closedMatch[4]!) - 1;
      } else if (rowBoundedMatch) {
        startCol = colToIndex(rowBoundedMatch[1]!);
        startRow = parseInt(rowBoundedMatch[2]!) - 1;
        endCol = Infinity;
        endRow = parseInt(rowBoundedMatch[3]!) - 1;
      } else if (colBoundedMatch) {
        startCol = colToIndex(colBoundedMatch[1]!);
        startRow = parseInt(colBoundedMatch[2]!) - 1;
        endCol = colToIndex(colBoundedMatch[3]!);
        endRow = Infinity;
      } else if (openBothMatch) {
        startCol = colToIndex(openBothMatch[1]!);
        startRow = parseInt(openBothMatch[2]!) - 1;
        endCol = Infinity;
        endRow = Infinity;
      } else {
        alert(
          "Invalid range format. Use format like [Workbook]Sheet!A1:C10 or A1:C10"
        );
        return;
      }

      // Normalize the range: ensure start is always before end
      if (endCol !== Infinity && startCol > endCol) {
        [startCol, endCol] = [endCol, startCol];
      }
      if (endRow !== Infinity && startRow > endRow) {
        [startRow, endRow] = [endRow, startRow];
      }

      if (editingCellStyle && editingCellStyle.workbookName === workbookName) {
        // Remove the old style
        engine.removeCellStyle(
          editingCellStyle.workbookName,
          editingCellStyle.index
        );
        // Add the new style
        engine.addCellStyle({
          areas: [{
            workbookName,
            sheetName,
            range: {
              start: { col: startCol, row: startRow },
              end: {
                col:
                  endCol === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endCol },
                row:
                  endRow === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endRow },
              },
            },
          }],
          style: {
            backgroundColor: newCellStyleBackgroundColor,
            color: newCellStyleColor,
            fontSize: parseInt(newCellStyleFontSize) || 12,
            bold: newCellStyleBold,
            italic: newCellStyleItalic,
            underline: newCellStyleUnderline,
          },
        });
        setEditingCellStyle(null);
        resetCellStyleForm();
      } else {
        engine.addCellStyle({
          areas: [{
            workbookName,
            sheetName,
            range: {
              start: { col: startCol, row: startRow },
              end: {
                col:
                  endCol === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endCol },
                row:
                  endRow === Infinity
                    ? { type: "infinity" as const, sign: "positive" as const }
                    : { type: "number" as const, value: endRow },
              },
            },
          }],
          style: {
            backgroundColor: newCellStyleBackgroundColor,
            color: newCellStyleColor,
            fontSize: parseInt(newCellStyleFontSize) || 12,
            bold: newCellStyleBold,
            italic: newCellStyleItalic,
            underline: newCellStyleUnderline,
          },
        });
      }

      markUnsavedChanges();
      resetCellStyleForm();
    } catch (error) {
      console.error("Failed to add cell style:", error);
      alert(
        `Failed to add cell style: ${
          error instanceof Error ? error.message : "Unknown error"
        }`
      );
    }
  }, [
    workbookGridItems,
    engine,
    markUnsavedChanges,
    newCellStyleArea,
    newCellStyleBackgroundColor,
    newCellStyleColor,
    newCellStyleFontSize,
    newCellStyleBold,
    newCellStyleItalic,
    newCellStyleUnderline,
    editingCellStyle,
    resetCellStyleForm,
  ]);

  // Delete cell style
  const deleteCellStyle = useCallback(
    (workbookName: string, index: number) => {
      try {
        engine.removeCellStyle(workbookName, index);
        markUnsavedChanges();
        // If we're editing this style, cancel editing
        if (
          editingCellStyle &&
          editingCellStyle.workbookName === workbookName &&
          editingCellStyle.index === index
        ) {
          setEditingCellStyle(null);
          resetCellStyleForm();
        }
      } catch (error) {
        console.error("Failed to delete cell style:", error);
        alert(
          `Failed to delete cell style: ${
            error instanceof Error ? error.message : "Unknown error"
          }`
        );
      }
    },
    [engine, markUnsavedChanges, editingCellStyle, resetCellStyleForm]
  );

  // Edit cell style
  const editCellStyle = useCallback(
    (workbookName: string, index: number) => {
      // Get style directly from engineState since index is global
      const style = engineState.cellStyles?.[index];
      if (!style) return;

      // Convert range to canonical format
      const colToLetter = (col: number): string => {
        let result = "";
        let c = col;
        while (c >= 0) {
          result = String.fromCharCode(65 + (c % 26)) + result;
          c = Math.floor(c / 26) - 1;
        }
        return result;
      };

      const startCol = colToLetter(style.areas[0]!.range.start.col);
      const startRow = style.areas[0]!.range.start.row + 1;
      const isColInfinity = style.areas[0]!.range.end.col.type === "infinity";
      const isRowInfinity = style.areas[0]!.range.end.row.type === "infinity";

      let cellRange: string;
      if (isColInfinity && isRowInfinity) {
        cellRange = `${startCol}${startRow}:INFINITY`;
      } else if (isColInfinity) {
        const endRow =
          style.areas[0]!.range.end.row.type === "number"
            ? style.areas[0]!.range.end.row.value + 1
            : startRow;
        cellRange = `${startCol}${startRow}:${endRow}`;
      } else if (isRowInfinity) {
        const endCol =
          style.areas[0]!.range.end.col.type === "number"
            ? colToLetter(style.areas[0]!.range.end.col.value)
            : startCol;
        cellRange = `${startCol}${startRow}:${endCol}`;
      } else {
        const endCol =
          style.areas[0]!.range.end.col.type === "number"
            ? colToLetter(style.areas[0]!.range.end.col.value)
            : startCol;
        const endRow =
          style.areas[0]!.range.end.row.type === "number"
            ? style.areas[0]!.range.end.row.value + 1
            : startRow;
        cellRange = `${startCol}${startRow}:${endCol}${endRow}`;
      }

      // Quote sheet name if needed
      const needsQuotes = /[ '!]/.test(style.areas[0]!.sheetName);
      const sheetRef = needsQuotes
        ? `'${style.areas[0]!.sheetName.replace(/'/g, "''")}'`
        : style.areas[0]!.sheetName;
      const rangeStr = `[${style.areas[0]!.workbookName}]${sheetRef}!${cellRange}`;

      setNewCellStyleArea(rangeStr);
      setNewCellStyleBackgroundColor(style.style.backgroundColor || "#FFFFFF");
      setNewCellStyleColor(style.style.color || "#000000");
      setNewCellStyleFontSize(String(style.style.fontSize || 12));
      setNewCellStyleBold(style.style.bold || false);
      setNewCellStyleItalic(style.style.italic || false);
      setNewCellStyleUnderline(style.style.underline || false);

      setEditingCellStyle({ workbookName, index });
    },
    [engineState]
  );

  // Handle selection change from spreadsheet
  const handleSelectionChange = useCallback(
    (
      selection: {
        workbookName: string;
        sheetName: string;
        area: {
          start: { col: number; row: number };
          end: {
            col: { type: string; value?: number } | number;
            row: { type: string; value?: number } | number;
          };
        };
      } | null
    ) => {
      // Always update the input when there's a selection
      // Only clear it when panel is visible and there's no selection
      if (!selection) {
        // Clear the input when there's no selection (only if panel is visible)
        if (showConditionalFormatting) {
          setNewStyleArea("");
        }
        // Also clear cell style area if needed
        setNewCellStyleArea("");
        return;
      }

      // Update the input with the selection (always, regardless of panel visibility)
      const colToLetter = (col: number): string => {
        let result = "";
        let c = col;
        while (c >= 0) {
          result = String.fromCharCode(65 + (c % 26)) + result;
          c = Math.floor(c / 26) - 1;
        }
        return result;
      };

      // Extract and normalize start/end coordinates
      let startColIdx = selection.area.start.col;
      let startRowIdx = selection.area.start.row;
      let endColIdx: number | "infinity";
      let endRowIdx: number | "infinity";

      // Extract end column
      if (typeof selection.area.end.col === "object") {
        if (selection.area.end.col.type === "infinity") {
          endColIdx = "infinity";
        } else if (
          selection.area.end.col.type === "number" &&
          selection.area.end.col.value !== undefined
        ) {
          endColIdx = selection.area.end.col.value;
        } else {
          endColIdx = startColIdx; // Fallback
        }
      } else {
        endColIdx = selection.area.end.col;
      }

      // Extract end row
      if (typeof selection.area.end.row === "object") {
        if (selection.area.end.row.type === "infinity") {
          endRowIdx = "infinity";
        } else if (
          selection.area.end.row.type === "number" &&
          selection.area.end.row.value !== undefined
        ) {
          endRowIdx = selection.area.end.row.value;
        } else {
          endRowIdx = startRowIdx; // Fallback
        }
      } else {
        endRowIdx = selection.area.end.row;
      }

      // Normalize: ensure start is always before end (swap if needed)
      if (
        endColIdx !== "infinity" &&
        typeof endColIdx === "number" &&
        startColIdx > endColIdx
      ) {
        [startColIdx, endColIdx] = [endColIdx, startColIdx];
      }
      if (
        endRowIdx !== "infinity" &&
        typeof endRowIdx === "number" &&
        startRowIdx > endRowIdx
      ) {
        [startRowIdx, endRowIdx] = [endRowIdx, startRowIdx];
      }

      // Convert to strings
      const startCol = colToLetter(startColIdx);
      const startRow = startRowIdx + 1; // Convert to 1-based

      let endColStr: string;
      let endRowStr: string;

      if (endColIdx === "infinity") {
        endColStr = "INFINITY";
      } else {
        endColStr = colToLetter(endColIdx);
      }

      if (endRowIdx === "infinity") {
        endRowStr = "INFINITY";
      } else {
        endRowStr = String(endRowIdx + 1); // Convert to 1-based
      }

      // Build canonical range format based on CANONICAL_RANGES.md
      let cellRange: string;
      if (endColStr === "INFINITY" && endRowStr === "INFINITY") {
        // Open both: A5:INFINITY
        cellRange = `${startCol}${startRow}:INFINITY`;
      } else if (endColStr === "INFINITY") {
        // Open→ (row-bounded): A5:10
        cellRange = `${startCol}${startRow}:${endRowStr}`;
      } else if (endRowStr === "INFINITY") {
        // Open↓ (col-bounded): A5:D
        cellRange = `${startCol}${startRow}:${endColStr}`;
      } else {
        // Closed rectangle: A5:D10
        cellRange = `${startCol}${startRow}:${endColStr}${endRowStr}`;
      }

      // Add workbook and sheet name in format: [workbook]'sheet'!range
      // Quote sheet name if it contains spaces or special characters
      const needsQuotes = /[ '!]/.test(selection.sheetName);
      const sheetRef = needsQuotes
        ? `'${selection.sheetName.replace(/'/g, "''")}'`
        : selection.sheetName;
      const rangeStr = `[${selection.workbookName}]${sheetRef}!${cellRange}`;

      setNewStyleArea(rangeStr);
      // Also update cell style area
      setNewCellStyleArea(rangeStr);
    },
    [showConditionalFormatting]
  );

  // Create WorkbookComponent for grid
  const WorkbookComponent = useCallback(
    ({ workbookName }: { workbookName: string }) => {
      const workbookItem = workbookGridItems.find(
        (item) => item.name === workbookName
      );
      if (!workbookItem) return null;

      // Get all sheets for this workbook
      const sheetNames = engine.getOrderedSheetNames(workbookName);

      // Add new sheet handler
      const addSheet = () => {
        try {
          const newSheet = engine.createSheet({
            workbookName,
          });

          // Switch to the new sheet
          updateWorkbookActiveSheet(workbookName, newSheet.name);
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
              clipboardManager={clipboardManager}
              key={`${workbookName}-${workbookItem.activeSheet}`}
              sheetName={workbookItem.activeSheet}
              workbookName={workbookName}
              engine={engine}
              verboseErrors={verboseErrors}
              onSelectionChange={handleSelectionChange}
            />
          </div>

          {/* Sheet Tabs at Bottom (Excel-style) */}
          <div
            className="h-14 bg-gray-50 border-t border-gray-200 flex items-center px-2 flex-shrink-0 w-full overflow-x-auto"
            style={{ scrollbarGutter: "stable" }}
          >
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
      handleSelectionChange,
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

  const numTables = Array.from(engineState.tables.values()).reduce(
    (sum, wb) => sum + wb.size,
    0
  );

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
                  <span
                    className="text-sm font-medium text-blue-800"
                    data-testid="current-filename"
                  >
                    {currentFileName}
                  </span>
                  {hasUnsavedChanges && (
                    <span
                      className="text-xs text-orange-600"
                      data-testid="unsaved-changes-indicator"
                    >
                      ●
                    </span>
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
                {/* Undo/Redo */}
                <div className="flex items-center gap-1 border-r border-gray-300 pr-2">
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={() => {
                      engine.undo();
                      setActionLog(engine.getActionLog());
                    }}
                    disabled={!engine.canUndo()}
                    title="Undo (Ctrl+Z)"
                    data-testid="undo-button"
                  >
                    <Undo className="h-4 w-4" />
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    className="border-gray-300 text-gray-700"
                    onClick={() => {
                      engine.redo();
                      setActionLog(engine.getActionLog());
                    }}
                    disabled={!engine.canRedo()}
                    title="Redo (Ctrl+Y)"
                    data-testid="redo-button"
                  >
                    <Redo className="h-4 w-4" />
                  </Button>
                </div>
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
                    title="Save to OPFS"
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
                  variant={showChangelog ? "default" : "outline"}
                  className={
                    showChangelog
                      ? "bg-blue-600 hover:bg-blue-700 text-white"
                      : "border-gray-300 text-gray-700"
                  }
                  onClick={() => {
                    setShowChangelog(!showChangelog);
                    if (!showChangelog) {
                      setActionLog(engine.getActionLog());
                    }
                  }}
                  data-testid="changelog-toggle"
                  title="Changelog"
                >
                  <History className="h-4 w-4 mr-1" />
                  Changelog ({actionLog.length})
                </Button>
                <Button
                  size="sm"
                  variant="outline"
                  className="border-gray-300 text-gray-700"
                  onClick={() => setShowNamedExpressions(!showNamedExpressions)}
                  data-testid="named-expressions-toggle"
                >
                  <Calculator className="h-4 w-4 mr-1" />
                  Expressions, Tables & Formatting
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

              {/* Conditional Formatting */}
              <div
                className="bg-white p-3 rounded border border-gray-200"
                data-testid="conditional-formatting-section"
              >
                <h3
                  className="text-sm font-semibold text-gray-800 mb-3"
                  data-testid="conditional-formatting-title"
                >
                  Conditional Formatting (
                  {engineState.conditionalStyles?.length || 0})
                </h3>
                <div className="space-y-3">
                  {/* Add Form */}
                  <div className="space-y-2 border-b border-gray-200 pb-2">
                    <div className="flex gap-2">
                      <Input
                        placeholder="Select a range in the spreadsheet..."
                        value={newStyleArea}
                        onChange={(e) => setNewStyleArea(e.target.value)}
                        className="text-xs flex-1"
                      />
                      <select
                        value={newStyleType}
                        onChange={(e) =>
                          setNewStyleType(
                            e.target.value as "formula" | "gradient"
                          )
                        }
                        className="text-xs border border-gray-300 rounded px-2"
                      >
                        <option value="formula">Formula</option>
                        <option value="gradient">Gradient</option>
                      </select>
                    </div>

                    {newStyleType === "formula" && (
                      <div className="space-y-2">
                        <Input
                          placeholder="Formula (e.g., ROW() > 4)"
                          value={newStyleFormula}
                          onChange={(e) => setNewStyleFormula(e.target.value)}
                          className="text-xs"
                        />
                        <div className="flex gap-2 items-center">
                          <label className="text-xs text-gray-600">
                            Color:
                          </label>
                          <input
                            type="color"
                            value={lchToHex({
                              l: parseFloat(newStyleColorL),
                              c: parseFloat(newStyleColorC),
                              h: parseFloat(newStyleColorH),
                            })}
                            onChange={(e) => {
                              const hex = e.target.value;
                              const lch = hexToLch(hex);
                              setNewStyleColorL(lch.l.toFixed(0));
                              setNewStyleColorC(lch.c.toFixed(0));
                              setNewStyleColorH(lch.h.toFixed(0));
                            }}
                            className="w-12 h-8 border border-gray-300 rounded cursor-pointer"
                          />
                          <span className="text-xs text-gray-500">
                            L:{newStyleColorL} C:{newStyleColorC} H:
                            {newStyleColorH}
                          </span>
                        </div>
                      </div>
                    )}

                    {newStyleType === "gradient" && (
                      <div className="space-y-2">
                        <select
                          value={newStyleGradientType}
                          onChange={(e) =>
                            setNewStyleGradientType(
                              e.target.value as "lowest_highest" | "number"
                            )
                          }
                          className="text-xs border border-gray-300 rounded px-2 w-full"
                        >
                          <option value="lowest_highest">
                            Lowest to Highest Value
                          </option>
                          <option value="number">Custom Min/Max</option>
                        </select>

                        {newStyleGradientType === "number" && (
                          <div className="flex gap-2">
                            <Input
                              placeholder="Min formula"
                              value={newStyleMinFormula}
                              onChange={(e) =>
                                setNewStyleMinFormula(e.target.value)
                              }
                              className="text-xs flex-1"
                            />
                            <Input
                              placeholder="Max formula"
                              value={newStyleMaxFormula}
                              onChange={(e) =>
                                setNewStyleMaxFormula(e.target.value)
                              }
                              className="text-xs flex-1"
                            />
                          </div>
                        )}

                        <div className="flex gap-2">
                          <div className="flex-1">
                            <label className="text-xs text-gray-600">
                              Min Color:
                            </label>
                            <div className="flex gap-1 items-center mt-1">
                              <input
                                type="color"
                                value={lchToHex({
                                  l: parseFloat(newStyleMinColorL),
                                  c: parseFloat(newStyleMinColorC),
                                  h: parseFloat(newStyleMinColorH),
                                })}
                                onChange={(e) => {
                                  const hex = e.target.value;
                                  const lch = hexToLch(hex);
                                  setNewStyleMinColorL(lch.l.toFixed(0));
                                  setNewStyleMinColorC(lch.c.toFixed(0));
                                  setNewStyleMinColorH(lch.h.toFixed(0));
                                }}
                                className="w-10 h-6 border border-gray-300 rounded cursor-pointer"
                              />
                            </div>
                          </div>
                          <div className="flex-1">
                            <label className="text-xs text-gray-600">
                              Max Color:
                            </label>
                            <div className="flex gap-1 items-center mt-1">
                              <input
                                type="color"
                                value={lchToHex({
                                  l: parseFloat(newStyleMaxColorL),
                                  c: parseFloat(newStyleMaxColorC),
                                  h: parseFloat(newStyleMaxColorH),
                                })}
                                onChange={(e) => {
                                  const hex = e.target.value;
                                  const lch = hexToLch(hex);
                                  setNewStyleMaxColorL(lch.l.toFixed(0));
                                  setNewStyleMaxColorC(lch.c.toFixed(0));
                                  setNewStyleMaxColorH(lch.h.toFixed(0));
                                }}
                                className="w-10 h-6 border border-gray-300 rounded cursor-pointer"
                              />
                            </div>
                          </div>
                        </div>
                      </div>
                    )}

                    <div className="flex gap-2">
                      {editingStyle && (
                        <Button
                          size="sm"
                          variant="outline"
                          onClick={() => {
                            setEditingStyle(null);
                            resetStyleForm();
                          }}
                          className="text-xs flex-1"
                        >
                          Cancel
                        </Button>
                      )}
                      <Button
                        size="sm"
                        onClick={addConditionalStyle}
                        disabled={!newStyleArea}
                        className="text-xs flex-1"
                      >
                        {editingStyle ? (
                          <>
                            <Edit2 className="h-3 w-3 mr-1" />
                            Update Rule
                          </>
                        ) : (
                          <>
                            <Plus className="h-3 w-3 mr-1" />
                            Add Rule
                          </>
                        )}
                      </Button>
                    </div>
                  </div>

                  {/* List of styles */}
                  <div
                    className="space-y-2 max-h-40 overflow-y-auto"
                    data-testid="conditional-formatting-list"
                  >
                    {(engineState.conditionalStyles || []).map(
                      (style, index) => {
                        const workbookName = style.areas[0]!.workbookName;
                        const colToLetter = (col: number): string => {
                          let result = "";
                          let c = col;
                          while (c >= 0) {
                            result =
                              String.fromCharCode(65 + (c % 26)) + result;
                            c = Math.floor(c / 26) - 1;
                          }
                          return result;
                        };

                        const rangeStr = `${colToLetter(
                          style.areas[0]!.range.start.col
                        )}${style.areas[0]!.range.start.row + 1}:${
                          style.areas[0]!.range.end.col.type === "infinity"
                            ? "∞"
                            : colToLetter(style.areas[0]!.range.end.col.value!)
                        }${
                          style.areas[0]!.range.end.row.type === "infinity"
                            ? "∞"
                            : style.areas[0]!.range.end.row.value! + 1
                        }`;

                        // Get color preview
                        const getColorPreview = () => {
                          if (style.condition.type === "formula") {
                            const color = style.condition.color;
                            return `hsl(${color.h}, ${color.c}%, ${color.l}%)`;
                          } else {
                            // Show gradient preview
                            const minColor = style.condition.min.color;
                            const maxColor = style.condition.max.color;
                            return `linear-gradient(90deg, hsl(${minColor.h}, ${minColor.c}%, ${minColor.l}%), hsl(${maxColor.h}, ${maxColor.c}%, ${maxColor.l}%))`;
                          }
                        };

                        return (
                          <div
                            key={`${workbookName}-${index}`}
                            className="flex items-center gap-2 bg-gray-50 p-2 rounded"
                            data-testid={`conditional-style-${index}`}
                          >
                            <div
                              className="w-6 h-6 rounded border border-gray-300 flex-shrink-0"
                              style={{ background: getColorPreview() }}
                              title="Color preview"
                            />
                            <div className="flex-1 min-w-0">
                              <div className="text-xs text-gray-800 truncate">
                                <span className="font-medium text-purple-600">
                                  {workbookName}
                                </span>
                                {" → "}
                                <span className="font-medium text-blue-600">
                                  {style.areas[0]!.sheetName}
                                </span>
                                {" • "}
                                <span className="font-medium">{rangeStr}</span>
                              </div>
                              <div className="text-xs text-gray-600 truncate">
                                {style.condition.type === "formula"
                                  ? `Formula: ${style.condition.formula}`
                                  : style.condition.min.type === "lowest_value"
                                  ? "Gradient: Min to Max"
                                  : `Gradient: ${
                                      style.condition.min.type === "number"
                                        ? style.condition.min.valueFormula
                                        : ""
                                    } to ${
                                      style.condition.max.type === "number"
                                        ? style.condition.max.valueFormula
                                        : ""
                                    }`}
                              </div>
                            </div>
                            <div className="flex items-center gap-1">
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-blue-500 hover:text-blue-700"
                                onClick={() =>
                                  editConditionalStyle(workbookName, index)
                                }
                                data-testid={`edit-conditional-style-${index}`}
                                title="Edit rule"
                              >
                                <Edit2 className="h-3 w-3" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                                onClick={() =>
                                  deleteConditionalStyle(workbookName, index)
                                }
                                data-testid={`delete-conditional-style-${index}`}
                                title="Delete rule"
                              >
                                <X className="h-3 w-3" />
                              </Button>
                            </div>
                          </div>
                        );
                      }
                    )}

                    {(!engineState.conditionalStyles ||
                      engineState.conditionalStyles.length === 0) && (
                      <p className="text-xs text-gray-500 italic">
                        No conditional formatting rules
                      </p>
                    )}
                  </div>
                </div>
              </div>

              {/* Cell Styles */}
              <div
                className="bg-white p-3 rounded border border-gray-200"
                data-testid="cell-styles-section"
              >
                <h3
                  className="text-sm font-semibold text-gray-800 mb-3"
                  data-testid="cell-styles-title"
                >
                  Cell Styles ({engineState.cellStyles?.length || 0})
                </h3>
                <div className="space-y-3">
                  {/* Add Form */}
                  <div className="space-y-2 border-b border-gray-200 pb-2">
                    <Input
                      placeholder="Select a range in the spreadsheet..."
                      value={newCellStyleArea}
                      onChange={(e) => setNewCellStyleArea(e.target.value)}
                      className="text-xs"
                    />
                    <div className="space-y-2">
                      <div className="flex gap-2 items-center">
                        <label className="text-xs text-gray-600">
                          Background:
                        </label>
                        <input
                          type="color"
                          value={newCellStyleBackgroundColor}
                          onChange={(e) =>
                            setNewCellStyleBackgroundColor(e.target.value)
                          }
                          className="w-12 h-8 border border-gray-300 rounded cursor-pointer"
                        />
                        <label className="text-xs text-gray-600">Text:</label>
                        <input
                          type="color"
                          value={newCellStyleColor}
                          onChange={(e) => setNewCellStyleColor(e.target.value)}
                          className="w-12 h-8 border border-gray-300 rounded cursor-pointer"
                        />
                      </div>
                      <div className="flex gap-2 items-center">
                        <label className="text-xs text-gray-600">
                          Font Size:
                        </label>
                        <input
                          type="number"
                          value={newCellStyleFontSize}
                          onChange={(e) =>
                            setNewCellStyleFontSize(e.target.value)
                          }
                          min="8"
                          max="72"
                          className="w-16 h-8 px-2 border border-gray-300 rounded"
                        />
                        <label className="flex items-center gap-1 text-xs">
                          <input
                            type="checkbox"
                            checked={newCellStyleBold}
                            onChange={(e) =>
                              setNewCellStyleBold(e.target.checked)
                            }
                            className="h-4 w-4"
                          />
                          <span className="font-bold">Bold</span>
                        </label>
                        <label className="flex items-center gap-1 text-xs">
                          <input
                            type="checkbox"
                            checked={newCellStyleItalic}
                            onChange={(e) =>
                              setNewCellStyleItalic(e.target.checked)
                            }
                            className="h-4 w-4"
                          />
                          <span className="italic">Italic</span>
                        </label>
                        <label className="flex items-center gap-1 text-xs">
                          <input
                            type="checkbox"
                            checked={newCellStyleUnderline}
                            onChange={(e) =>
                              setNewCellStyleUnderline(e.target.checked)
                            }
                            className="h-4 w-4"
                          />
                          <span className="underline">Underline</span>
                        </label>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <Button
                        onClick={addCellStyle}
                        disabled={!newCellStyleArea}
                        size="sm"
                        className="text-xs flex-1"
                      >
                        {editingCellStyle ? (
                          <>
                            <Edit2 className="h-3 w-3 mr-1" /> Update Style
                          </>
                        ) : (
                          <>
                            <Plus className="h-3 w-3 mr-1" /> Add Style
                          </>
                        )}
                      </Button>
                      <Button
                        onClick={() => {
                          if (!newCellStyleArea) return;

                          // Parse range
                          let workbookName: string;
                          let sheetName: string;
                          let rangeStr: string;

                          const fullRangeMatch = newCellStyleArea.match(
                            /^\[([^\]]+)\](?:'([^']+(?:''[^']*)*)'|([^!]+))!(.+)$/
                          );

                          if (fullRangeMatch) {
                            workbookName = fullRangeMatch[1]!;
                            sheetName = fullRangeMatch[2]
                              ? fullRangeMatch[2].replace(/''/g, "'")
                              : fullRangeMatch[3]!;
                            rangeStr = fullRangeMatch[4]!;
                          } else {
                            workbookName = workbookGridItems[0]?.name!;
                            sheetName = workbookGridItems[0]?.activeSheet!;
                            rangeStr = newCellStyleArea;
                          }

                          // Parse range coordinates (reuse parsing logic)
                          const colToIndex = (col: string): number => {
                            let result = 0;
                            for (let i = 0; i < col.length; i++) {
                              result = result * 26 + (col.charCodeAt(i) - 64);
                            }
                            return result - 1;
                          };

                          let startCol: number, startRow: number;
                          let endCol: number, endRow: number;

                          const closedMatch = rangeStr.match(
                            /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/
                          );
                          const rowBoundedMatch = rangeStr.match(
                            /^([A-Z]+)(\d+):(\d+)$/
                          );
                          const colBoundedMatch = rangeStr.match(
                            /^([A-Z]+)(\d+):([A-Z]+)$/
                          );
                          const openBothMatch = rangeStr.match(
                            /^([A-Z]+)(\d+):INFINITY$/
                          );

                          if (closedMatch) {
                            startCol = colToIndex(closedMatch[1]!);
                            startRow = parseInt(closedMatch[2]!) - 1;
                            endCol = colToIndex(closedMatch[3]!);
                            endRow = parseInt(closedMatch[4]!) - 1;
                          } else if (rowBoundedMatch) {
                            startCol = colToIndex(rowBoundedMatch[1]!);
                            startRow = parseInt(rowBoundedMatch[2]!) - 1;
                            endCol = Infinity;
                            endRow = parseInt(rowBoundedMatch[3]!) - 1;
                          } else if (colBoundedMatch) {
                            startCol = colToIndex(colBoundedMatch[1]!);
                            startRow = parseInt(colBoundedMatch[2]!) - 1;
                            endCol = colToIndex(colBoundedMatch[3]!);
                            endRow = Infinity;
                          } else if (openBothMatch) {
                            startCol = colToIndex(openBothMatch[1]!);
                            startRow = parseInt(openBothMatch[2]!) - 1;
                            endCol = Infinity;
                            endRow = Infinity;
                          } else {
                            alert("Invalid range format");
                            return;
                          }

                          engine.clearCellStyles({
                            workbookName,
                            sheetName,
                            range: {
                              start: { col: startCol, row: startRow },
                              end: {
                                col:
                                  endCol === Infinity
                                    ? {
                                        type: "infinity" as const,
                                        sign: "positive" as const,
                                      }
                                    : {
                                        type: "number" as const,
                                        value: endCol,
                                      },
                                row:
                                  endRow === Infinity
                                    ? {
                                        type: "infinity" as const,
                                        sign: "positive" as const,
                                      }
                                    : {
                                        type: "number" as const,
                                        value: endRow,
                                      },
                              },
                            },
                          });

                          markUnsavedChanges();
                        }}
                        disabled={!newCellStyleArea}
                        size="sm"
                        variant="destructive"
                        className="text-xs"
                        title="Clear all styles in selected range"
                      >
                        <Trash2 className="h-3 w-3 mr-1" />
                        Clear Range
                      </Button>
                      {editingCellStyle && (
                        <Button
                          onClick={resetCellStyleForm}
                          size="sm"
                          variant="outline"
                          className="text-xs"
                        >
                          Cancel
                        </Button>
                      )}
                    </div>
                  </div>

                  {/* List of styles */}
                  <div
                    className="space-y-2 max-h-40 overflow-y-auto"
                    data-testid="cell-styles-list"
                  >
                    {(engineState.cellStyles || []).map((style, index) => {
                      const workbookName = style.areas[0]!.workbookName;
                      const colToLetter = (col: number): string => {
                        let result = "";
                        let c = col;
                        while (c >= 0) {
                          result = String.fromCharCode(65 + (c % 26)) + result;
                          c = Math.floor(c / 26) - 1;
                        }
                        return result;
                      };

                      const rangeStr = `${colToLetter(
                        style.areas[0]!.range.start.col
                      )}${style.areas[0]!.range.start.row + 1}:${
                        style.areas[0]!.range.end.col.type === "infinity"
                          ? "∞"
                          : colToLetter(style.areas[0]!.range.end.col.value!)
                      }${
                        style.areas[0]!.range.end.row.type === "infinity"
                          ? "∞"
                          : style.areas[0]!.range.end.row.value! + 1
                      }`;

                      return (
                        <div
                          key={`${workbookName}-${index}`}
                          className="flex items-center gap-2 bg-gray-50 p-2 rounded"
                          data-testid={`cell-style-${index}`}
                        >
                          <div
                            className="w-6 h-6 rounded border border-gray-300 flex-shrink-0"
                            style={{
                              backgroundColor: style.style.backgroundColor,
                              color: style.style.color,
                            }}
                            title="Style preview"
                          >
                            <div className="w-full h-full flex items-center justify-center text-xs">
                              A
                            </div>
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="text-xs text-gray-800 truncate">
                              <span className="font-medium text-purple-600">
                                {workbookName}
                              </span>
                              {" → "}
                              <span className="font-medium text-blue-600">
                                {style.areas[0]!.sheetName}
                              </span>
                              {" • "}
                              <span className="font-medium">{rangeStr}</span>
                            </div>
                            <div className="text-xs text-gray-600 truncate">
                              BG: {style.style.backgroundColor} • Text:{" "}
                              {style.style.color}
                            </div>
                          </div>
                          <div className="flex items-center gap-1">
                            <Button
                              size="sm"
                              variant="ghost"
                              className="h-6 w-6 p-0 text-blue-500 hover:text-blue-700"
                              onClick={() => editCellStyle(workbookName, index)}
                              data-testid={`edit-cell-style-${index}`}
                              title="Edit style"
                            >
                              <Edit2 className="h-3 w-3" />
                            </Button>
                            <Button
                              size="sm"
                              variant="ghost"
                              className="h-6 w-6 p-0 text-red-500 hover:text-red-700"
                              onClick={() =>
                                deleteCellStyle(workbookName, index)
                              }
                              data-testid={`delete-cell-style-${index}`}
                              title="Delete style"
                            >
                              <X className="h-3 w-3" />
                            </Button>
                          </div>
                        </div>
                      );
                    })}

                    {(!engineState.cellStyles ||
                      engineState.cellStyles.length === 0) && (
                      <p className="text-xs text-gray-500 italic">
                        No cell styles
                      </p>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Changelog Sidebar */}
      {showChangelog && (
        <div
          className="border-b border-gray-200 bg-gray-50 p-4"
          data-testid="changelog-panel"
        >
          <div className="bg-white p-3 rounded border border-gray-200">
            <div className="flex items-center justify-between mb-3">
              <h3 className="text-sm font-semibold text-gray-800">
                Changelog ({actionLog.length} actions)
              </h3>
              <Button
                size="sm"
                variant="ghost"
                className="h-6 w-6 p-0"
                onClick={() => setShowChangelog(false)}
                title="Close Changelog"
              >
                <X className="h-4 w-4" />
              </Button>
            </div>
            <div className="space-y-2 max-h-60 overflow-y-auto">
              {actionLog.length === 0 ? (
                <p className="text-xs text-gray-500 italic">
                  No actions recorded yet
                </p>
              ) : (
                actionLog
                  .slice()
                  .reverse()
                  .map((action, index) => {
                    const actionIndex = actionLog.length - index - 1;
                    const formattedAction = formatAction(action);
                    return (
                      <div
                        key={`${action.type}-${actionIndex}-${action.timestamp}`}
                        className="flex items-start gap-2 bg-gray-50 p-2 rounded text-xs"
                        data-testid={`changelog-action-${actionIndex}`}
                      >
                        <div className="flex-shrink-0 w-8 text-right text-gray-500 font-mono">
                          {actionIndex + 1}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 mb-1">
                            <span className="font-medium text-gray-800">
                              {formattedAction.type}
                            </span>
                            {action.timestamp && (
                              <span className="text-gray-500">
                                {new Date(action.timestamp).toLocaleTimeString()}
                              </span>
                            )}
                          </div>
                          {formattedAction.description && (
                            <div className="text-gray-600 text-xs">
                              {formattedAction.description}
                            </div>
                          )}
                          {formattedAction.details && (
                            <div className="text-gray-500 text-xs mt-1 font-mono">
                              {formattedAction.details}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })
              )}
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
