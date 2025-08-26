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

export function useSerializedSheet(
  engine: FormulaEngine,
  sheetName: string
): SerializedSheet {
  const [serialized, setSerialized] = useState<SerializedSheet>({
    sheet: new Map(engine.getSheetSerialized(sheetName)),
    namedExpressions: new Map(engine.getNamedExpressionsSerialized(sheetName)),
  });

  React.useEffect(() => {
    return engine.onCellsUpdate(sheetName, () => {
      const sheet = new Map(engine.getSheetSerialized(sheetName));
      const namedExpressions = new Map(
        engine.getNamedExpressionsSerialized(sheetName)
      );
      setSerialized({
        sheet,
        namedExpressions,
      });
    });
  }, [engine, sheetName]);

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
    return engine.on("tables-updated", (tables) => {
      setTables(new Map(tables));
    });
  }, [engine]);

  return tables;
}