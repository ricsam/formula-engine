/**
 * React hooks for FormulaEngine integration
 */

import React, { useState } from "react";
import type { FormulaEngine } from "../core/engine";
import type { CellValue } from "../core/types";

export function useSerializedSheet(
  engine: FormulaEngine,
  sheetId: number
): Map<string, CellValue> {
  const [serialized, setSerialized] = useState<Map<string, CellValue>>(
    new Map()
  );

  React.useEffect(() => {
    return engine.onCellsUpdate(sheetId, (events) => {
      if (events.length > 0) {
        setSerialized(engine.getSheetSerialized(sheetId));
      }
    });
  }, [engine, sheetId]);

  return serialized;
}
