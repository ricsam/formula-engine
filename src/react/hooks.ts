/**
 * React hooks for FormulaEngine integration
 */

import React, { useState } from "react";
import type { FormulaEngine } from "../core/engine";
import type { SerializedCellValue } from "../core/types";

export function useSerializedSheet(
  engine: FormulaEngine,
  sheetName: string
): Map<string, SerializedCellValue> {
  const [serialized, setSerialized] = useState<
    Map<string, SerializedCellValue>
  >(() => {
    return new Map(engine.getSheetSerialized(sheetName));
  });

  React.useEffect(() => {
    return engine.onCellsUpdate(sheetName, () => {
      setSerialized(new Map(engine.getSheetSerialized(sheetName)));
    });
  }, [engine, sheetName]);

  return serialized;
}
