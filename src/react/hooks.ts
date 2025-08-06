/**
 * React hooks for FormulaEngine integration
 */

import { useState, useEffect, useCallback, useMemo } from 'react';
import type { FormulaEngine } from '../core/engine';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../core/types';
import type {
  UseSpreadsheetOptions,
  UseCellOptions,
  UseSpreadsheetRangeOptions,
  SpreadsheetHookResult,
  CellHookResult,
  SpreadsheetRangeHookResult
} from './types';

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
  
  const [spreadsheet, setSpreadsheet] = useState<Map<string, CellValue>>(new Map());
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
      setError(err instanceof Error ? err : new Error('Failed to get sheet ID'));
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
      setError(err instanceof Error ? err : new Error('Failed to get sheet contents'));
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
    const unsubscribeCellChanges = engine.on('cell-changed', (event) => {
      if (event.address.sheet === sheetId) {
        debouncedUpdate();
      }
    });

    // Subscribe to sheet operations that affect this sheet
    const unsubscribeSheetRemoved = engine.on('sheet-removed', (event) => {
      if (event.sheetId === sheetId) {
        setSpreadsheet(new Map());
        setError(new Error(`Sheet "${sheetName}" was removed`));
      }
    });

    const unsubscribeSheetRenamed = engine.on('sheet-renamed', (event) => {
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

/**
 * Custom hook to subscribe to a single cell value
 * Returns the current value and auto-updates when it changes
 */
export function useCell(
  engine: FormulaEngine,
  sheetName: string,
  cellAddress: string, // e.g., "A1", "B5"
  options: UseCellOptions = {}
): CellHookResult {
  const { autoUpdate = true, debounceMs = 0 } = options;
  
  const [value, setValue] = useState<CellValue>(() => {
    try {
      const sheetId = engine.getSheetId(sheetName);
      const address = engine.simpleCellAddressFromString(cellAddress, sheetId);
      return engine.getCellValue(address);
    } catch {
      return undefined;
    }
  });
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<Error | null>(null);

  // Parse address once when dependencies change
  const address = useMemo(() => {
    try {
      if (engine.doesSheetExist(sheetName)) {
        const sheetId = engine.getSheetId(sheetName);
        return engine.simpleCellAddressFromString(cellAddress, sheetId);
      }
      return null;
    } catch (err) {
      setError(err instanceof Error ? err : new Error('Failed to parse cell address'));
      return null;
    }
  }, [engine, sheetName, cellAddress]);

  // Function to update cell value
  const updateCell = useCallback(() => {
    if (address === null) {
      setValue(undefined);
      setIsLoading(false);
      return;
    }

    try {
      const newValue = engine.getCellValue(address);
      setValue(newValue);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err : new Error('Failed to get cell value'));
      setValue(undefined);
    } finally {
      setIsLoading(false);
    }
  }, [engine, address]);

  // Debounced update function
  const debouncedUpdate = useMemo(() => {
    if (debounceMs === 0) {
      return updateCell;
    }

    let timeoutId: NodeJS.Timeout;
    return () => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(updateCell, debounceMs);
    };
  }, [updateCell, debounceMs]);

  useEffect(() => {
    // Initial load
    updateCell();

    if (!autoUpdate || address === null) {
      return;
    }

    // Subscribe to changes affecting this specific cell
    const unsubscribeCellChanges = engine.on('cell-changed', (event) => {
      if (
        event.address.sheet === address.sheet &&
        event.address.col === address.col &&
        event.address.row === address.row
      ) {
        debouncedUpdate();
      }
    });

    // Subscribe to dependency changes that might affect this cell
    const unsubscribeDependencyUpdates = engine.on('dependency-updated', (event) => {
      const isAffected = event.affectedCells.some(addr => 
        addr.sheet === address.sheet &&
        addr.col === address.col &&
        addr.row === address.row
      );
      
      if (isAffected) {
        debouncedUpdate();
      }
    });

    // Subscribe to sheet removal
    const unsubscribeSheetRemoved = engine.on('sheet-removed', (event) => {
      if (event.sheetId === address.sheet) {
        setValue(undefined);
        setError(new Error(`Sheet "${sheetName}" was removed`));
      }
    });

    return () => {
      unsubscribeCellChanges();
      unsubscribeDependencyUpdates();
      unsubscribeSheetRemoved();
    };
  }, [engine, address, sheetName, autoUpdate, debouncedUpdate]);

  return { value, isLoading, error };
}

/**
 * Custom hook to subscribe to a range of cells
 * Returns a Map of cell addresses to values for the specified range
 */
export function useSpreadsheetRange(
  engine: FormulaEngine,
  sheetName: string,
  range: string, // e.g., "A1:C10"
  options: UseSpreadsheetRangeOptions = {}
): SpreadsheetRangeHookResult {
  const { autoUpdate = true, debounceMs = 0 } = options;
  
  const [rangeData, setRangeData] = useState<Map<string, CellValue>>(new Map());
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<Error | null>(null);

  // Parse range once when dependencies change
  const { sheetId, cellRange } = useMemo(() => {
    try {
      if (engine.doesSheetExist(sheetName)) {
        const id = engine.getSheetId(sheetName);
        const parsedRange = engine.simpleCellRangeFromString(range, id);
        return { sheetId: id, cellRange: parsedRange };
      }
      return { sheetId: null, cellRange: null };
    } catch (err) {
      setError(err instanceof Error ? err : new Error('Failed to parse range'));
      return { sheetId: null, cellRange: null };
    }
  }, [engine, sheetName, range]);

  // Function to update range data
  const updateRange = useCallback(() => {
    if (sheetId === null || cellRange === null) {
      setRangeData(new Map());
      setIsLoading(false);
      return;
    }

    try {
      const allContents = engine.getSheetContents(sheetId);
      const filtered = new Map<string, CellValue>();
      
      // Filter to only include cells in the specified range
      for (const [addressStr, value] of allContents) {
        try {
          const cellAddr = engine.simpleCellAddressFromString(addressStr, sheetId);
          if (isInRange(cellAddr, cellRange)) {
            filtered.set(addressStr, value);
          }
        } catch {
          // Skip invalid addresses
          continue;
        }
      }
      
      setRangeData(filtered);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err : new Error('Failed to get range contents'));
      setRangeData(new Map());
    } finally {
      setIsLoading(false);
    }
  }, [engine, sheetId, cellRange]);

  // Debounced update function
  const debouncedUpdate = useMemo(() => {
    if (debounceMs === 0) {
      return updateRange;
    }

    let timeoutId: NodeJS.Timeout;
    return () => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(updateRange, debounceMs);
    };
  }, [updateRange, debounceMs]);

  useEffect(() => {
    // Initial load
    updateRange();

    if (!autoUpdate || sheetId === null || cellRange === null) {
      return;
    }

    // Subscribe to cell changes in this range
    const unsubscribeCellChanges = engine.on('cell-changed', (event) => {
      if (event.address.sheet === sheetId && isInRange(event.address, cellRange)) {
        debouncedUpdate();
      }
    });

    // Subscribe to sheet removal
    const unsubscribeSheetRemoved = engine.on('sheet-removed', (event) => {
      if (event.sheetId === sheetId) {
        setRangeData(new Map());
        setError(new Error(`Sheet "${sheetName}" was removed`));
      }
    });

    return () => {
      unsubscribeCellChanges();
      unsubscribeSheetRemoved();
    };
  }, [engine, sheetId, cellRange, sheetName, autoUpdate, debouncedUpdate]);

  return { rangeData, isLoading, error };
}

/**
 * Utility function to check if an address is within a range
 */
function isInRange(address: SimpleCellAddress, range: SimpleCellRange): boolean {
  return (
    address.sheet === range.start.sheet &&
    address.col >= range.start.col && 
    address.col <= range.end.col &&
    address.row >= range.start.row && 
    address.row <= range.end.row
  );
}

/**
 * Hook to monitor FormulaEngine events
 * Useful for debugging or implementing custom event handling
 */
export function useFormulaEngineEvents(
  engine: FormulaEngine,
  events: {
    onCellChanged?: (event: { address: SimpleCellAddress; oldValue: CellValue; newValue: CellValue }) => void;
    onSheetAdded?: (event: { sheetId: number; sheetName: string }) => void;
    onSheetRemoved?: (event: { sheetId: number; sheetName: string }) => void;
    onSheetRenamed?: (event: { sheetId: number; oldName: string; newName: string }) => void;
    onFormulaCalculated?: (event: { address: SimpleCellAddress; formula: string; result: CellValue }) => void;
    onDependencyUpdated?: (event: { affectedCells: SimpleCellAddress[] }) => void;
  }
): void {
  useEffect(() => {
    const unsubscribers: (() => void)[] = [];

    if (events.onCellChanged) {
      unsubscribers.push(engine.on('cell-changed', events.onCellChanged));
    }

    if (events.onSheetAdded) {
      unsubscribers.push(engine.on('sheet-added', events.onSheetAdded));
    }

    if (events.onSheetRemoved) {
      unsubscribers.push(engine.on('sheet-removed', events.onSheetRemoved));
    }

    if (events.onSheetRenamed) {
      unsubscribers.push(engine.on('sheet-renamed', events.onSheetRenamed));
    }

    if (events.onFormulaCalculated) {
      unsubscribers.push(engine.on('formula-calculated', events.onFormulaCalculated));
    }

    if (events.onDependencyUpdated) {
      unsubscribers.push(engine.on('dependency-updated', events.onDependencyUpdated));
    }

    return () => {
      unsubscribers.forEach(unsubscribe => unsubscribe());
    };
  }, [engine, events]);
}
