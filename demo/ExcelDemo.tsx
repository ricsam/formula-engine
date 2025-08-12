import { useState, useCallback, useMemo } from "react";
import { FormulaEngine } from "../src/core/engine";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Plus, X, Edit2, Check, X as Cancel } from "lucide-react";
import { SpreadsheetWithFormulaBar } from "./components/SpreadsheetWithFormulaBar";

interface SheetTab {
  name: string;
}

const createEngine = () => {
  const engine = FormulaEngine.buildEmpty();

  // Create first sheet with sample data
  const sheetName = engine.addSheet("Sheet1").name;

  return { engine, sheetName };
};

export function ExcelDemo() {
  const {
    engine,
    sheetName: initialSheetName,
  } = useMemo(() => createEngine(), []);

  const [sheets, setSheets] = useState<SheetTab[]>([
    {
      name: initialSheetName,
    },
  ]);

  const [activeSheet, setActiveSheet] = useState(initialSheetName);
  const [editingSheet, setEditingSheet] = useState<string | null>(null);
  const [editingName, setEditingName] = useState("");

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
  }, [sheets.length, engine]);

  // Delete sheet
  const deleteSheet = useCallback(
    (sheetName: string) => {
      if (sheets.length <= 1) return; // Don't delete the last sheet

      engine.removeSheet(sheetName);

      setSheets((prev) => {
        const newSheets = prev.filter((sheet) => sheet.name !== sheetName);

        // If we deleted the active sheet, switch to the first remaining sheet
        if (sheetName === activeSheet && newSheets.length > 0) {
          setActiveSheet(newSheets[0]!.name);
        }

        return newSheets;
      });
    },
    [sheets.length, activeSheet, engine]
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
      engine.renameSheet(editingSheet, editingName.trim());

      setSheets((prev) =>
        prev.map((sheet) =>
          sheet.name === editingSheet
            ? { ...sheet, name: editingName.trim() }
            : sheet
        )
      );
    }

    setEditingSheet(null);
    setEditingName("");
  }, [editingSheet, editingName, engine]);

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
              <span>
                Active Sheet:{" "}
                <strong>
                  {sheets.find((s) => s.name === activeSheet)?.name}
                </strong>
              </span>
              <span>Total Sheets: {sheets.length}</span>
            </div>
          </div>
        </div>
      </div>

      {/* Main spreadsheet area with formula bar */}
      <div className="flex-1 overflow-hidden">
        <SpreadsheetWithFormulaBar
          key={activeSheet} // Re-mount component when sheet changes
          sheetName={activeSheet}
          engine={engine}
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
