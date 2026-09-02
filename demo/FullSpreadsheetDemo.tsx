import type {
  FormulaAnalysis,
  FormulaReference,
  FormulaReferenceInsertion,
  FormulaReferenceTarget,
} from "@ricsam/formula-engine-editor";
import { beginFormulaReferenceInsertion } from "@ricsam/formula-engine-editor";
import {
  FormulaEngine,
  indexToColumn,
  type CellAddress,
  type SerializedCellValue,
  type SpreadsheetRange,
} from "@ricsam/formula-engine";
import {
  FormulaEditor,
  type FormulaEditorHandle,
} from "@ricsam/formula-engine-editor/react";
import {
  coerceCellInput,
  FormulaWorkbook,
  getCellReference,
  WorkbookSelectionManager,
} from "@ricsam/react-spreadsheets";
import "@ricsam/react-spreadsheets/styles.css";
import type {
  SelectionManager,
  SMArea,
} from "@ricsam/selection-manager";
import type { editor as MonacoEditor } from "monaco-editor";
import React, {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import "./FullSpreadsheetDemo.css";

const WORKBOOK = "Studio";
const FORECAST_SHEET = "Forecast";
const ASSUMPTIONS_SHEET = "Assumptions";

const sheet = (sheetName: string) => ({ workbookName: WORKBOOK, sheetName });

const finiteRange = (
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number,
): SpreadsheetRange => ({
  start: { col: startCol, row: startRow },
  end: {
    col: { type: "number", value: endCol },
    row: { type: "number", value: endRow },
  },
});

function populateEngine(engine: FormulaEngine) {
  engine.setSheetContent(
    sheet(FORECAST_SHEET),
    new Map<string, SerializedCellValue>([
      ["A1", "PRODUCT"],
      ["B1", "UNITS"],
      ["C1", "UNIT PRICE"],
      ["D1", "REVENUE"],
      ["E1", "NET REVENUE"],
      ["A2", "Starter"],
      ["B2", 120],
      ["C2", 29],
      ["D2", "=B2*C2"],
      ["E2", "=D2*(1-Assumptions!B3)"],
      ["A3", "Pro"],
      ["B3", 65],
      ["C3", 99],
      ["D3", "=B3*C3"],
      ["E3", "=D3*(1-Assumptions!B3)"],
      ["A4", "Team"],
      ["B4", 28],
      ["C4", 249],
      ["D4", "=B4*C4"],
      ["E4", "=D4*(1-Assumptions!B3)"],
      ["A5", "Enterprise"],
      ["B5", 8],
      ["C5", 1200],
      ["D5", "=B5*C5"],
      ["E5", "=D5*(1-Assumptions!B3)"],
      ["A7", "CURRENT PERIOD"],
      ["B7", "=SUM(B2:B5)"],
      ["C7", "=AVERAGE(C2:C5)"],
      ["D7", "=SUM(D2:D5)"],
      ["E7", "=SUM(E2:E5)"],
      ["A9", "NEXT PERIOD"],
      ["B9", "Growth"],
      ["C9", "=Assumptions!B2"],
      ["D9", "Projected"],
      ["E9", "=E7*C9"],
      ["A11", "Try the editor"],
      ["B11", "Place the caret over B2, C2, or Assumptions!B3"],
      ["D40", "Off-screen navigation target"],
    ]),
  );

  engine.setSheetContent(
    sheet(ASSUMPTIONS_SHEET),
    new Map<string, SerializedCellValue>([
      ["A1", "ASSUMPTION"],
      ["B1", "VALUE"],
      ["C1", "DESCRIPTION"],
      ["A2", "Growth multiplier"],
      ["B2", 1.18],
      ["C2", "Used by Forecast!E9"],
      ["A3", "Platform fee"],
      ["B3", 0.07],
      ["C3", "Used by Forecast!E2:E5"],
      ["A4", "Target margin"],
      ["B4", 0.72],
      ["C4", "Planning target"],
      ["A6", "After fees"],
      ["B6", "=1-B3"],
      ["C6", "Derived locally"],
    ]),
  );
}

function createDemoEngine() {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(WORKBOOK);
  engine.addSheet(sheet(FORECAST_SHEET));
  engine.addSheet(sheet(ASSUMPTIONS_SHEET));
  populateEngine(engine);

  engine.addCellStyle({
    areas: [
      { ...sheet(FORECAST_SHEET), range: finiteRange(0, 0, 4, 0) },
      { ...sheet(ASSUMPTIONS_SHEET), range: finiteRange(0, 0, 2, 0) },
    ],
    style: {
      bold: true,
      backgroundColor: "#172033",
      color: "#ffffff",
    },
  });
  engine.addCellStyle({
    areas: [{ ...sheet(FORECAST_SHEET), range: finiteRange(0, 6, 4, 6) }],
    style: {
      bold: true,
      backgroundColor: "#eef2ff",
      color: "#3730a3",
      borderColor: "#c7d2fe",
    },
  });
  engine.addCellStyle({
    areas: [{ ...sheet(FORECAST_SHEET), range: finiteRange(0, 8, 4, 8) }],
    style: {
      bold: true,
      backgroundColor: "#ecfdf5",
      color: "#047857",
    },
  });
  engine.clearUndoRedoHistory();
  return engine;
}

function readRawCell(engine: FormulaEngine, address: CellAddress): string {
  const ref = getCellReference(address);
  const raw = engine
    .getSheet({ workbookName: address.workbookName, sheetName: address.sheetName })
    ?.content.get(ref);
  return raw === undefined ? "" : String(raw);
}

type PhysicalTarget = {
  workbookName: string;
  sheetName: string;
  range: SpreadsheetRange;
};

function physicalTarget(target: FormulaReferenceTarget): PhysicalTarget | undefined {
  if (target.type === "cell") {
    return {
      workbookName: target.address.workbookName,
      sheetName: target.address.sheetName,
      range: finiteRange(
        target.address.colIndex,
        target.address.rowIndex,
        target.address.colIndex,
        target.address.rowIndex,
      ),
    };
  }
  if (target.type === "range") return target.address;
  if (target.type === "table") {
    return {
      workbookName: target.workbookName,
      sheetName: target.sheetName,
      range: target.range,
    };
  }
  return undefined;
}

function resolvedTargets(reference: FormulaReference | undefined): PhysicalTarget[] {
  if (!reference || reference.resolution.status !== "resolved") return [];
  return reference.resolution.targets
    .map(physicalTarget)
    .filter((target): target is PhysicalTarget => target !== undefined);
}

function isCellInRange(
  rowIndex: number,
  colIndex: number,
  range: SpreadsheetRange,
) {
  const rowInRange =
    rowIndex >= range.start.row &&
    (range.end.row.type === "infinity" || rowIndex <= range.end.row.value);
  const colInRange =
    colIndex >= range.start.col &&
    (range.end.col.type === "infinity" || colIndex <= range.end.col.value);
  return rowInRange && colInRange;
}

function formatRange(range: SpreadsheetRange): string {
  const start = getCellReference({
    colIndex: range.start.col,
    rowIndex: range.start.row,
  });
  if (range.end.col.type === "infinity" || range.end.row.type === "infinity") {
    return `${start}:∞`;
  }
  const end = getCellReference({
    colIndex: range.end.col.value,
    rowIndex: range.end.row.value,
  });
  return start === end ? start : `${start}:${end}`;
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)
    ? sheetName
    : `'${sheetName.replaceAll("'", "''")}'`;
}

/** Formats a grid selection as a formula reference relative to the edited cell. */
function formatPickedReference(
  area: SMArea,
  pickedSheetName: string,
  origin: CellAddress,
): string {
  const startRow = area.start.row;
  const startCol = area.start.col;
  const endRow = area.end.row;
  const endCol = area.end.col;

  let localReference: string;
  if (endRow.type === "infinity" && endCol.type === "infinity") {
    localReference = `${getCellReference({
      rowIndex: startRow,
      colIndex: startCol,
    })}:INFINITY`;
  } else if (endRow.type === "infinity") {
    const lastCol = endCol.type === "number" ? endCol.value : startCol;
    const first = indexToColumn(Math.min(startCol, lastCol));
    const last = indexToColumn(Math.max(startCol, lastCol));
    localReference = startRow === 0 ? `${first}:${last}` : `${first}${startRow + 1}:${last}`;
  } else if (endCol.type === "infinity") {
    const firstRow = Math.min(startRow, endRow.value) + 1;
    const lastRow = Math.max(startRow, endRow.value) + 1;
    localReference =
      startCol === 0
        ? `${firstRow}:${lastRow}`
        : `${indexToColumn(startCol)}${firstRow}:${lastRow}`;
  } else {
    const firstRow = Math.min(startRow, endRow.value);
    const lastRow = Math.max(startRow, endRow.value);
    const firstCol = Math.min(startCol, endCol.value);
    const lastCol = Math.max(startCol, endCol.value);
    const first = getCellReference({ rowIndex: firstRow, colIndex: firstCol });
    const last = getCellReference({ rowIndex: lastRow, colIndex: lastCol });
    localReference = first === last ? first : `${first}:${last}`;
  }

  const sheetQualifier =
    pickedSheetName === origin.sheetName ? "" : `${quoteSheetName(pickedSheetName)}!`;
  return `${sheetQualifier}${localReference}`;
}

function singleCellArea(address: CellAddress): SMArea {
  return {
    start: { row: address.rowIndex, col: address.colIndex },
    end: {
      row: { type: "number", value: address.rowIndex },
      col: { type: "number", value: address.colIndex },
    },
  };
}

function formatTarget(target: FormulaReferenceTarget): string {
  if (target.type === "cell") {
    return `${target.address.sheetName}!${getCellReference(target.address)}`;
  }
  if (target.type === "range") {
    return `${target.address.sheetName}!${formatRange(target.address.range)}`;
  }
  if (target.type === "table") return `${target.sheetName}!${target.tableName}`;
  return target.name;
}

function EditorIcon({ name }: { name: "apply" | "reset" | "spark" }) {
  if (name === "apply") {
    return <svg viewBox="0 0 20 20" aria-hidden="true"><path d="m4 10 4 4 8-8" /></svg>;
  }
  if (name === "reset") {
    return <svg viewBox="0 0 20 20" aria-hidden="true"><path d="M4.5 6.5A6 6 0 1 1 4 13M4.5 6.5V2.8m0 3.7H8" /></svg>;
  }
  return (
    <svg viewBox="0 0 20 20" aria-hidden="true">
      <path d="m10 2-1 3.2A4 4 0 0 1 6.2 8L3 9l3.2 1A4 4 0 0 1 9 12.8l1 3.2 1-3.2a4 4 0 0 1 2.8-2.8L17 9l-3.2-1A4 4 0 0 1 11 5.2L10 2Z" />
    </svg>
  );
}

export function FullSpreadsheetDemo() {
  const engine = useMemo(createDemoEngine, []);
  const workbookSelectionManager = useMemo(() => new WorkbookSelectionManager(), []);
  const initialAddress = useMemo<CellAddress>(
    () => ({
      workbookName: WORKBOOK,
      sheetName: FORECAST_SHEET,
      colIndex: 3,
      rowIndex: 1,
    }),
    [],
  );
  const [selectedAddress, setSelectedAddress] = useState<CellAddress>(initialAddress);
  const [savedFormula, setSavedFormula] = useState(() => readRawCell(engine, initialAddress));
  const [draft, setDraft] = useState(savedFormula);
  const [analysis, setAnalysis] = useState<FormulaAnalysis>();
  const [activeReference, setActiveReference] = useState<FormulaReference>();
  const [saveState, setSaveState] = useState<"saved" | "editing" | "error">("saved");
  const [saveMessage, setSaveMessage] = useState("Saved");
  const [isPickingReference, setIsPickingReference] = useState(false);
  const [referenceSheetOverride, setReferenceSheetOverride] = useState<string>();
  const [, setRevision] = useState(0);
  const formulaEditorRef = useRef<FormulaEditorHandle>(null);
  const currentSheetSelectionManagerRef = useRef<SelectionManager | undefined>(
    undefined,
  );
  const referenceInsertionRef = useRef<FormulaReferenceInsertion | undefined>(undefined);
  const lastInsertedReferenceSpanRef = useRef<
    FormulaReferenceInsertion["span"] | undefined
  >(undefined);
  const isFormulaEditingRef = useRef(false);
  const isClosingFormulaEditingRef = useRef(false);
  const suppressNextFormulaEditActivationRef = useRef(false);
  const lastCellBySheet = useRef(
    new Map<string, { colIndex: number; rowIndex: number }>([
      [FORECAST_SHEET, { colIndex: 3, rowIndex: 1 }],
      [ASSUMPTIONS_SHEET, { colIndex: 1, rowIndex: 1 }],
    ]),
  );
  const draftRef = useRef(draft);
  const selectedAddressRef = useRef(selectedAddress);
  const isDirty = draft !== savedFormula;

  draftRef.current = draft;
  selectedAddressRef.current = selectedAddress;

  useEffect(
    () =>
      engine.onUpdate(() => {
        setRevision((value) => value + 1);
        formulaEditorRef.current?.refresh();
      }),
    [engine],
  );

  const targets = useMemo(() => resolvedTargets(activeReference), [activeReference]);
  const previewTarget = useMemo(() => {
    if (targets.length === 0) return undefined;
    return targets.find((target) => target.sheetName === selectedAddress.sheetName) ?? targets[0];
  }, [selectedAddress.sheetName, targets]);
  const visibleSheet =
    referenceSheetOverride ?? previewTarget?.sheetName ?? selectedAddress.sheetName;
  const isCrossSheetPreview = visibleSheet !== selectedAddress.sheetName;

  const loadAddress = useCallback(
    (address: CellAddress) => {
      isClosingFormulaEditingRef.current = true;
      try {
        referenceInsertionRef.current?.cancel();
      } finally {
        isClosingFormulaEditingRef.current = false;
      }
      referenceInsertionRef.current = undefined;
      lastInsertedReferenceSpanRef.current = undefined;
      currentSheetSelectionManagerRef.current?.endReferenceSelection();
      isFormulaEditingRef.current = false;
      setIsPickingReference(false);
      setReferenceSheetOverride(undefined);
      lastCellBySheet.current.set(address.sheetName, {
        colIndex: address.colIndex,
        rowIndex: address.rowIndex,
      });
      const nextFormula = readRawCell(engine, address);
      formulaEditorRef.current?.getEditor()?.setPosition({
        lineNumber: 1,
        column: 1,
      });
      setSelectedAddress(address);
      setSavedFormula(nextFormula);
      setDraft(nextFormula);
      setAnalysis(undefined);
      setActiveReference(undefined);
      setSaveState("saved");
      setSaveMessage("Saved");
    },
    [engine],
  );

  useEffect(
    () =>
      workbookSelectionManager.onSelectionChange((selections) => {
        const selection = selections[selections.length - 1];
        if (!selection) return;
        const current = selectedAddressRef.current;
        const { workbookName, sheetName, range } = selection;
        if (
          current.workbookName === workbookName &&
          current.sheetName === sheetName &&
          current.rowIndex === range.start.row &&
          current.colIndex === range.start.col
        ) return;

        loadAddress({
          workbookName,
          sheetName,
          rowIndex: range.start.row,
          colIndex: range.start.col,
        });
      }),
    [loadAddress, workbookSelectionManager],
  );

  const handleGridSelection = useCallback(
    (selectionManager: SelectionManager) => {
      currentSheetSelectionManagerRef.current = selectionManager;
      if (isFormulaEditingRef.current) {
        selectionManager.beginReferenceSelection({
          editedRange: singleCellArea(selectedAddressRef.current),
        });
      }

      const cleanup = selectionManager.listenToReferenceSelection((event) => {
        if (event.phase === "cancel") {
          isClosingFormulaEditingRef.current = true;
          try {
            referenceInsertionRef.current?.cancel();
          } finally {
            isClosingFormulaEditingRef.current = false;
          }
          referenceInsertionRef.current = undefined;
          lastInsertedReferenceSpanRef.current = undefined;
          isFormulaEditingRef.current = false;
          selectionManager.endReferenceSelection();
          setIsPickingReference(false);
          setReferenceSheetOverride(undefined);
          return;
        }

        let insertion = referenceInsertionRef.current;
        if (!insertion) {
          const editor = formulaEditorRef.current?.getEditor();
          if (
            !isFormulaEditingRef.current ||
            !editor?.getModel()?.getValue().trimStart().startsWith("=")
          ) {
            return;
          }

          const selection = editor.getSelection();
          const activeReferenceSpan = selection?.isEmpty()
            ? formulaEditorRef.current?.getActiveReference()?.span
            : undefined;
          const replaceSpan =
            lastInsertedReferenceSpanRef.current ?? activeReferenceSpan;

          insertion = beginFormulaReferenceInsertion(
            editor,
            replaceSpan === undefined ? undefined : { replaceSpan },
          );
          if (!insertion) return;
          referenceInsertionRef.current = insertion;
          setIsPickingReference(true);
        }

        lastInsertedReferenceSpanRef.current = insertion.update(
          formatPickedReference(event.range, visibleSheet, selectedAddressRef.current),
        );

        if (event.phase === "commit") {
          insertion.finish();
          referenceInsertionRef.current = undefined;
          setIsPickingReference(false);
        }
      });

      return () => {
        cleanup();
        selectionManager.endReferenceSelection();
        if (currentSheetSelectionManagerRef.current === selectionManager) {
          currentSheetSelectionManagerRef.current = undefined;
        }
      };
    },
    [visibleSheet],
  );

  const applyDraft = useCallback(() => {
    const address = selectedAddressRef.current;
    try {
      const content = coerceCellInput(draftRef.current, engine.getCellDataType(address));
      engine.setCellContent(address, content);
      const normalized = readRawCell(engine, address);
      setSavedFormula(normalized);
      setDraft(normalized);
      setSaveState("saved");
      setSaveMessage("Applied to the sheet");
      isClosingFormulaEditingRef.current = true;
      try {
        referenceInsertionRef.current?.finish();
      } finally {
        isClosingFormulaEditingRef.current = false;
      }
      referenceInsertionRef.current = undefined;
      currentSheetSelectionManagerRef.current?.endReferenceSelection();
      isFormulaEditingRef.current = false;
      lastInsertedReferenceSpanRef.current = undefined;
      setIsPickingReference(false);
      setReferenceSheetOverride(undefined);
      formulaEditorRef.current?.refresh();
    } catch (error) {
      setSaveState("error");
      setSaveMessage(error instanceof Error ? error.message : "Could not apply formula");
    }
  }, [engine]);

  const revertDraft = useCallback(() => {
    const current = readRawCell(engine, selectedAddressRef.current);
    isClosingFormulaEditingRef.current = true;
    try {
      referenceInsertionRef.current?.cancel();
    } finally {
      isClosingFormulaEditingRef.current = false;
    }
    referenceInsertionRef.current = undefined;
    currentSheetSelectionManagerRef.current?.endReferenceSelection();
    suppressNextFormulaEditActivationRef.current = draftRef.current !== current;
    isFormulaEditingRef.current = false;
    lastInsertedReferenceSpanRef.current = undefined;
    setIsPickingReference(false);
    setReferenceSheetOverride(undefined);
    setSavedFormula(current);
    setDraft(current);
    setSaveState("saved");
    setSaveMessage("Changes reverted");
    setActiveReference(undefined);
  }, [engine]);

  const applyDraftRef = useRef(applyDraft);
  const revertDraftRef = useRef(revertDraft);
  applyDraftRef.current = applyDraft;
  revertDraftRef.current = revertDraft;

  const handleEditorMount = useCallback(
    (editor: MonacoEditor.IStandaloneCodeEditor, monaco: typeof import("monaco-editor")) => {
      const activateFormulaEditing = () => {
        if (isClosingFormulaEditingRef.current) return;
        const isFormula = editor.getValue().trimStart().startsWith("=");
        isFormulaEditingRef.current = isFormula;
        const selectionManager = currentSheetSelectionManagerRef.current;
        if (isFormula) {
          if (selectionManager?.selectionMode !== "reference") {
            selectionManager?.beginReferenceSelection({
              editedRange: singleCellArea(selectedAddressRef.current),
            });
          }
        } else {
          referenceInsertionRef.current?.cancel();
          referenceInsertionRef.current = undefined;
          lastInsertedReferenceSpanRef.current = undefined;
          selectionManager?.endReferenceSelection();
          setIsPickingReference(false);
          setReferenceSheetOverride(undefined);
        }
      };

      const clearCommittedReferenceSelection = () => {
        if (isClosingFormulaEditingRef.current) return;
        const selectionManager = currentSheetSelectionManagerRef.current;
        if (
          !referenceInsertionRef.current &&
          isFormulaEditingRef.current &&
          selectionManager?.selectionMode === "reference"
        ) {
          selectionManager.beginReferenceSelection({
            editedRange: singleCellArea(selectedAddressRef.current),
          });
        }
      };

      editor.onDidFocusEditorText(activateFormulaEditing);
      editor.onMouseDown(activateFormulaEditing);
      editor.onDidChangeCursorSelection(() => {
        if (!referenceInsertionRef.current) {
          lastInsertedReferenceSpanRef.current = undefined;
          clearCommittedReferenceSelection();
        }
      });
      editor.onDidChangeModelContent(() => {
        if (!referenceInsertionRef.current) {
          lastInsertedReferenceSpanRef.current = undefined;
          clearCommittedReferenceSelection();
        }
        if (suppressNextFormulaEditActivationRef.current) {
          suppressNextFormulaEditActivationRef.current = false;
          return;
        }
        if (editor.hasTextFocus()) activateFormulaEditing();
      });
      editor.addCommand(monaco.KeyCode.Enter, () => applyDraftRef.current());
      editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter, () => applyDraftRef.current());
      editor.addCommand(monaco.KeyCode.Escape, () => revertDraftRef.current());
    },
    [],
  );

  const handleSheetChange = useCallback(
    (sheetName: string) => {
      if (isFormulaEditingRef.current) {
        setReferenceSheetOverride(sheetName);
        return;
      }
      const previous = lastCellBySheet.current.get(sheetName) ?? { colIndex: 0, rowIndex: 0 };
      loadAddress({ workbookName: WORKBOOK, sheetName, ...previous });
    },
    [loadAddress],
  );

  const handleReset = useCallback(() => {
    populateEngine(engine);
    engine.clearUndoRedoHistory();
    loadAddress(initialAddress);
    setSaveMessage("Demo reset");
  }, [engine, initialAddress, loadAddress]);

  const activeTargetLabel = useMemo(() => {
    if (!activeReference) return "Move the caret onto a reference";
    if (activeReference.resolution.status === "unresolved") {
      return `Unresolved · ${activeReference.resolution.reason.replaceAll("-", " ")}`;
    }
    if (activeReference.resolution.status === "dynamic") return "Dynamic reference";
    return activeReference.resolution.targets.map(formatTarget).join(", ");
  }, [activeReference]);

  const evaluatedSelectedValue = engine.getCellValue(selectedAddress, true);
  const selectedCellRef = getCellReference(selectedAddress);
  const selectedCellLabel = `${selectedAddress.sheetName}!${selectedCellRef}`;
  const diagnostics = analysis?.diagnostics ?? [];
  const errorCount = diagnostics.filter((item) => item.severity === "error").length;

  // The workbook selection is intentionally never mutated for caret previews.
  // Reference targets are a presentation-only cell style layered over the grid.
  const selectionInitialState = useMemo<{ selections: SMArea[] }>(
    () => ({
      selections: isCrossSheetPreview
        ? []
        : [{
            start: { row: selectedAddress.rowIndex, col: selectedAddress.colIndex },
            end: {
              row: { type: "number" as const, value: selectedAddress.rowIndex },
              col: { type: "number" as const, value: selectedAddress.colIndex },
            },
          }],
    }),
    [isCrossSheetPreview, selectedAddress.colIndex, selectedAddress.rowIndex],
  );

  return (
    <div className="formula-studio" data-testid="formula-studio">
      <header className="formula-studio__header">
        <div>
          <span className="formula-studio__eyebrow">Formula language tooling</span>
          <h1>Write formulas with the grid in view.</h1>
          <p>
            Monaco consumes <code>@ricsam/formula-engine-editor</code>. Type a formula,
            then click or drag across the grid to insert a cell or range at the cursor. Put
            the caret on any reference to preview its resolved target.
          </p>
        </div>
        <button type="button" className="formula-studio__reset" onClick={handleReset}>
          <EditorIcon name="reset" /> Reset demo
        </button>
      </header>

      <section className="formula-studio__workbench" aria-label="Formula editor demo">
        <div className="formula-studio__toolbar">
          <div className="formula-studio__cell-pill">
            <span>Editing</span>
            <strong data-testid="selected-cell-address">{selectedCellLabel}</strong>
          </div>
          <div className="formula-studio__value">
            <span>Calculated value</span>
            <strong data-testid="selected-cell-value">
              {evaluatedSelectedValue === undefined ? "—" : String(evaluatedSelectedValue)}
            </strong>
          </div>
          <div className={`formula-studio__save-state formula-studio__save-state--${saveState}`} role="status">
            <i /> {isDirty ? "Unapplied changes" : saveMessage}
          </div>
        </div>

        <div className="formula-studio__editor-panel">
          <div className="formula-studio__editor-heading">
            <div>
              <span className="formula-studio__fx">ƒx</span>
              <div>
                <strong>Formula editor</strong>
                <small>Click or drag cells to insert · Enter to apply · Esc to revert</small>
              </div>
            </div>
            <div className="formula-studio__editor-actions">
              <span className={errorCount > 0 ? "has-errors" : ""} data-testid="formula-diagnostics">
                {errorCount > 0 ? `${errorCount} ${errorCount === 1 ? "error" : "errors"}` : "Syntax valid"}
              </span>
              <button type="button" onClick={revertDraft} disabled={!isDirty}>Revert</button>
              <button
                type="button"
                className="formula-studio__apply"
                data-testid="apply-formula"
                onClick={applyDraft}
                disabled={!isDirty || errorCount > 0}
              >
                <EditorIcon name="apply" /> Apply
              </button>
            </div>
          </div>

          <div className="formula-studio__monaco">
            <FormulaEditor
              ref={formulaEditorRef}
              engine={engine}
              origin={selectedAddress}
              value={draft}
              height="112px"
              theme="formula-studio-theme"
              testId="formula-editor"
              beforeMount={(monaco) => {
                monaco.editor.defineTheme("formula-studio-theme", {
                  base: "vs",
                  inherit: true,
                  rules: [
                    { token: "function", foreground: "7C3AED", fontStyle: "bold" },
                    { token: "variable.formulaCellReference", foreground: "EA580C" },
                    { token: "namespace", foreground: "0369A1" },
                    { token: "number", foreground: "047857" },
                    { token: "string", foreground: "BE123C" },
                    { token: "operator", foreground: "475569" },
                  ],
                  colors: {
                    "editor.background": "#FBFCFE",
                    "editor.foreground": "#172033",
                    "editor.lineHighlightBackground": "#F1F5F900",
                    "editorCursor.foreground": "#EA580C",
                    "editor.selectionBackground": "#FED7AA88",
                    "editorError.foreground": "#DC2626",
                    "editorWarning.foreground": "#D97706",
                  },
                });
              }}
              onMount={handleEditorMount}
              onChange={(value, nextAnalysis) => {
                setDraft(value);
                setAnalysis(nextAnalysis);
                setSaveState("editing");
                setSaveMessage("Editing");
              }}
              onAnalysisChange={({ analysis: nextAnalysis }) => setAnalysis(nextAnalysis)}
              onActiveReferenceChange={(reference) => setActiveReference(reference)}
              options={{
                automaticLayout: true,
                contextmenu: false,
                cursorBlinking: "smooth",
                fontFamily: "'JetBrains Mono', 'SFMono-Regular', Consolas, monospace",
                fontLigatures: true,
                fontSize: 15,
                folding: false,
                glyphMargin: false,
                lineDecorationsWidth: 0,
                lineNumbers: "off",
                minimap: { enabled: false },
                overviewRulerBorder: false,
                overviewRulerLanes: 0,
                padding: { top: 21, bottom: 18 },
                renderLineHighlight: "none",
                scrollBeyondLastLine: false,
                scrollbar: { horizontal: "hidden", vertical: "hidden", alwaysConsumeMouseWheel: false },
                "semanticHighlighting.enabled": true,
                wordWrap: "on",
              }}
            />
          </div>

          <div
            className={`formula-studio__reference ${activeReference || isPickingReference ? "is-active" : ""}`}
            data-testid="active-reference"
          >
            <span className="formula-studio__reference-icon"><EditorIcon name="spark" /></span>
            <div>
              <small>
                {isPickingReference
                  ? "Inserting grid selection"
                  : isCrossSheetPreview
                    ? "Cross-sheet preview"
                    : "Caret target"}
              </small>
              <strong>{isPickingReference ? "Drag to resize the reference" : activeTargetLabel}</strong>
            </div>
            {activeReference && <code>{draft.slice(activeReference.span.start, activeReference.span.end)}</code>}
          </div>
        </div>

        <div className="formula-studio__grid-panel">
          <div className="formula-studio__grid-heading">
            <div>
              <strong data-testid="visible-sheet">{visibleSheet}</strong>
              <span>{isCrossSheetPreview ? `Previewing a reference from ${selectedAddress.sheetName}` : "Live workbook"}</span>
            </div>
            <div className="formula-studio__legend">
              <span><i className="selection" /> Selection</span>
              <span><i className="reference" /> Caret reference</span>
            </div>
          </div>
          <div className="formula-studio__grid" data-testid="formula-workbook">
            <FormulaWorkbook
              engine={engine}
              workbookName={WORKBOOK}
              activeSheet={visibleSheet}
              onActiveSheetChange={handleSheetChange}
              selectionManager={workbookSelectionManager}
              selection={{
                initialState: selectionInitialState,
                effects: handleGridSelection,
              }}
              isSelected
              customCellStyle={(cell, internalStyle) => {
                const highlighted = targets.some(
                  (target) =>
                    target.workbookName === WORKBOOK &&
                    target.sheetName === visibleSheet &&
                    isCellInRange(cell.rowIndex, cell.colIndex, target.range),
                );
                if (!highlighted) return internalStyle;
                return {
                  ...internalStyle,
                  backgroundColor: "#fff7ed",
                  outline: "2px solid #f97316",
                  outlineOffset: "-2px",
                  color: "#9a3412",
                  fontWeight: 700,
                  zIndex: 3,
                };
              }}
            />
          </div>
        </div>
      </section>

      <footer className="formula-studio__footer">
        <span><kbd>1</kbd> Select a formula cell</span><i />
        <span><kbd>2</kbd> Edit with semantic highlighting</span><i />
        <span><kbd>3</kbd> Click or drag cells to insert references</span><i />
        <span><kbd>4</kbd> Press Enter to recalculate</span>
      </footer>
    </div>
  );
}
