/**
 * React hooks for FormulaEngine integration
 */

import { useState, useEffect, useCallback, useMemo } from "react";
import type { FormulaEngine } from "../core/engine";
import type {
  CellValue,
  SimpleCellAddress,
  SimpleCellRange,
} from "../core/types";
import type {
  UseSpreadsheetOptions,
  UseCellOptions,
  UseSpreadsheetRangeOptions,
  SpreadsheetHookResult,
  CellHookResult,
  SpreadsheetRangeHookResult,
} from "./types";
import React from "react";

/**
 * Custom hook to subscribe to an entire spreadsheet sheet
 * Returns a reactive Map of all cell values
 */
export function useSpreadsheet(
  engine: FormulaEngine,
  sheetName: string,
  options: UseSpreadsheetOptions = {}
): SpreadsheetHookResult {
  const { autoUpdate = true, debounceMs = 0 } = options;

  const [spreadsheet, setSpreadsheet] = useState<Map<string, CellValue>>(
    new Map()
  );
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<Error | null>(null);

  // Get sheet ID once when dependencies change
  const sheetId = useMemo(() => {
    try {
      if (engine.doesSheetExist(sheetName)) {
        return engine.getSheetId(sheetName);
      }
      return null;
    } catch (err) {
      setError(
        err instanceof Error ? err : new Error("Failed to get sheet ID")
      );
      return null;
    }
  }, [engine, sheetName]);

  // Function to update spreadsheet data
  const updateSpreadsheet = useCallback(() => {
    if (sheetId === null) {
      setSpreadsheet(new Map());
      setIsLoading(false);
      return;
    }

    try {
      const contents = engine.getSheetContents(sheetId);
      setSpreadsheet(new Map(contents));
      setError(null);
    } catch (err) {
      setError(
        err instanceof Error ? err : new Error("Failed to get sheet contents")
      );
      setSpreadsheet(new Map());
    } finally {
      setIsLoading(false);
    }
  }, [engine, sheetId]);

  // Debounced update function
  const debouncedUpdate = useMemo(() => {
    if (debounceMs === 0) {
      return updateSpreadsheet;
    }

    let timeoutId: NodeJS.Timeout;
    return () => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(updateSpreadsheet, debounceMs);
    };
  }, [updateSpreadsheet, debounceMs]);

  useEffect(() => {
    // Initial load
    updateSpreadsheet();

    if (!autoUpdate || sheetId === null) {
      return;
    }

    // Subscribe to cell changes
    const unsubscribeCellChanges = engine.on("cell-changed", (event) => {
      if (event.address.sheet === sheetId) {
        debouncedUpdate();
      }
    });

    // Subscribe to sheet operations that affect this sheet
    const unsubscribeSheetRemoved = engine.on("sheet-removed", (event) => {
      if (event.sheetId === sheetId) {
        setSpreadsheet(new Map());
        setError(new Error(`Sheet "${sheetName}" was removed`));
      }
    });

    const unsubscribeSheetRenamed = engine.on("sheet-renamed", (event) => {
      if (event.sheetId === sheetId && event.newName !== sheetName) {
        // Sheet was renamed, but we're still tracking by the old name
        // This might indicate a stale reference
        debouncedUpdate();
      }
    });

    return () => {
      unsubscribeCellChanges();
      unsubscribeSheetRemoved();
      unsubscribeSheetRenamed();
    };
  }, [engine, sheetId, sheetName, autoUpdate, debouncedUpdate]);

  return { spreadsheet, isLoading, error };
}

export function useSerializedSheet(
  engine: FormulaEngine,
  sheetId: number
): Map<string, CellValue> {
  const [serialized, setSerialized] = useState<Map<string, CellValue>>(
    new Map()
  );

  React.useEffect(() => {
    return engine.subscribe("cells-changed", (events) => {
      const changes = events.filter((event) => event.address.sheet === sheetId);
      if (changes.length > 0) {
        setSerialized(engine.getSheetSerialized(sheetId));
      }
    });
  }, [engine, sheetId]);

  return serialized;
}
