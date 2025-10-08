/**
 * React hooks for FormulaEngine integration
 */

import React, { useState } from "react";
import type { FormulaEngine } from "../core/engine";

export function useEngine(
  engine: FormulaEngine
): ReturnType<typeof FormulaEngine.prototype.getState> {
  const [state, setState] = useState<
    ReturnType<typeof FormulaEngine.prototype.getState>
  >(() => engine.getState());

  React.useEffect(() => {
    return engine.onUpdate(() => {
      setState(engine.getState());
    });
  }, [engine]);

  return state;
}
