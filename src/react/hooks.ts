/**
 * React hooks for FormulaEngine integration
 */

import React, { useState } from "react";
import type { FormulaEngine } from "../core/engine";
import type {
  NamedExpression,
  SerializedCellValue,
  TableDefinition,
} from "../core/types";

export type SerializedSheet = {
  sheet: Map<string, SerializedCellValue>;
  namedExpressions: Map<string, NamedExpression>;
};

export function useSerializedSheet(opts: {
  engine: FormulaEngine;
  sheetName: string;
  workbookName: string;
}): SerializedSheet {
  const { engine, sheetName, workbookName } = opts;
  const [serialized, setSerialized] = useState<SerializedSheet>({
    sheet: new Map(engine.getSheetSerialized(opts)),
    namedExpressions: new Map(engine.getSheetExpressionsSerialized(opts)),
  });

  React.useEffect(() => {
    return engine.onCellsUpdate(opts, () => {
      const sheet = new Map(engine.getSheetSerialized(opts));
      const namedExpressions = new Map(
        engine.getSheetExpressionsSerialized(opts)
      );
      setSerialized({
        sheet,
        namedExpressions,
      });
    });
  }, [engine, sheetName, workbookName]);

  return serialized;
}

export function useGlobalNamedExpressions(
  engine: FormulaEngine
): Map<string, NamedExpression> {
  const [namedExpressions, setNamedExpressions] = useState<
    Map<string, NamedExpression>
  >(engine.getGlobalNamedExpressionsSerialized());

  React.useEffect(() => {
    return engine.on(
      "global-named-expressions-updated",
      (globalNamedExpressions) => {
        setNamedExpressions(new Map(globalNamedExpressions));
      }
    );
  }, [engine]);

  return namedExpressions;
}

export function useTables(engine: FormulaEngine): Map<string, TableDefinition> {
  const [tables, setTables] = useState<Map<string, TableDefinition>>(
    engine.getTablesSerialized()
  );

  React.useEffect(() => {
    return engine.on("tables-updated", (updatedTables) => {
      setTables(new Map(updatedTables));
    });
  }, [engine]);

  return tables;
}
