import type {
  FormulaEngineEvents,
  NamedExpression,
} from "../types";
import { renameNamedExpressionInFormula } from "../named-expression-renamer";

export class NamedExpressionManager {
  private scopedNamedExpressions: Map<string, Map<string, NamedExpression>> = new Map();
  private globalNamedExpressions: Map<string, NamedExpression> = new Map();
  private eventEmitter?: {
    emit<K extends keyof FormulaEngineEvents>(
      event: K,
      data: FormulaEngineEvents[K]
    ): void;
  };

  constructor(eventEmitter?: {
    emit<K extends keyof FormulaEngineEvents>(
      event: K,
      data: FormulaEngineEvents[K]
    ): void;
  }) {
    this.eventEmitter = eventEmitter;
  }

  getScopedNamedExpressions(): Map<string, Map<string, NamedExpression>> {
    return this.scopedNamedExpressions;
  }

  getGlobalNamedExpressions(): Map<string, NamedExpression> {
    return this.globalNamedExpressions;
  }

  addNamedExpression({
    expression,
    expressionName,
    sheetName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
  }): void {
    if (!sheetName) {
      this.globalNamedExpressions.set(expressionName, {
        name: expressionName,
        expression,
      });
      this.eventEmitter?.emit(
        "global-named-expressions-updated",
        this.globalNamedExpressions
      );
    } else {
      let scopedNamedExpressions = this.scopedNamedExpressions.get(sheetName);
      if (!scopedNamedExpressions) {
        scopedNamedExpressions = new Map();
        this.scopedNamedExpressions.set(sheetName, scopedNamedExpressions);
      }

      scopedNamedExpressions.set(expressionName, {
        name: expressionName,
        expression,
      });
    }
  }

  removeNamedExpression({
    expressionName,
    sheetName,
  }: {
    expressionName: string;
    sheetName?: string;
  }): boolean {
    let found = false;

    if (!sheetName) {
      // Remove from global named expressions
      found = this.globalNamedExpressions.delete(expressionName);
      if (found) {
        this.eventEmitter?.emit(
          "global-named-expressions-updated",
          this.globalNamedExpressions
        );
      }
    } else {
      // Remove from sheet-scoped named expressions
      const scopedNamedExpressions = this.scopedNamedExpressions.get(sheetName);
      if (scopedNamedExpressions) {
        found = scopedNamedExpressions.delete(expressionName);
      }
    }

    return found;
  }

  updateNamedExpression({
    expression,
    expressionName,
    sheetName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
  }): void {
    // Check if the named expression exists
    const exists = sheetName
      ? this.scopedNamedExpressions.get(sheetName)?.has(expressionName)
      : this.globalNamedExpressions.has(expressionName);

    if (!exists) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Update is the same as add for existing expressions
    this.addNamedExpression({ expression, expressionName, sheetName });
  }

  renameNamedExpression({
    expressionName,
    sheetName,
    newName,
  }: {
    expressionName: string;
    sheetName?: string;
    newName: string;
  }): boolean {
    // Check if the named expression exists
    const targetMap = sheetName
      ? this.scopedNamedExpressions.get(sheetName)
      : this.globalNamedExpressions;

    if (!targetMap || !targetMap.has(expressionName)) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Check if the new name already exists
    if (targetMap.has(newName)) {
      throw new Error(`Named expression '${newName}' already exists`);
    }

    // Get the expression to rename
    const namedExpression = targetMap.get(expressionName)!;

    // Update the name and re-add with new name
    const updatedExpression = { ...namedExpression, name: newName };
    targetMap.set(newName, updatedExpression);
    targetMap.delete(expressionName);

    this.eventEmitter?.emit("global-named-expressions-updated", this.globalNamedExpressions);

    return true;
  }

  updateFormulasForNamedExpressionRename(
    oldName: string,
    newName: string,
    updateCallback: (formula: string) => string = (formula) =>
      renameNamedExpressionInFormula(formula, oldName, newName)
  ): void {
    // Update global named expressions that reference this named expression
    this.globalNamedExpressions.forEach((namedExpr, name) => {
      if (name !== oldName) {
        // Don't update the expression we're renaming
        const updatedExpression = updateCallback(namedExpr.expression);

        if (updatedExpression !== namedExpr.expression) {
          this.globalNamedExpressions.set(name, {
            ...namedExpr,
            expression: updatedExpression,
          });
        }
      }
    });

    // Update scoped named expressions that reference this named expression
    this.scopedNamedExpressions.forEach((namedExpressionsMap, sheetName) => {
      namedExpressionsMap.forEach((namedExpr, name) => {
        if (name !== oldName) {
          // Don't update the expression we're renaming
          const updatedExpression = updateCallback(namedExpr.expression);

          if (updatedExpression !== namedExpr.expression) {
            namedExpressionsMap.set(name, {
              ...namedExpr,
              expression: updatedExpression,
            });
          }
        }
      });
    });
  }

  getNamedExpressionsSerialized(
    sheetName: string
  ): Map<string, NamedExpression> {
    return this.scopedNamedExpressions.get(sheetName) ?? new Map();
  }

  getGlobalNamedExpressionsSerialized(): Map<string, NamedExpression> {
    return this.globalNamedExpressions;
  }



  removeSheetNamedExpressions(sheetName: string): void {
    this.scopedNamedExpressions.delete(sheetName);
  }

  renameSheetNamedExpressions(oldName: string, newName: string): void {
    const namedExpressions = this.scopedNamedExpressions.get(oldName);
    if (namedExpressions) {
      this.scopedNamedExpressions.set(newName, namedExpressions);
      this.scopedNamedExpressions.delete(oldName);
    }
  }

  /**
   * Replace all global named expressions (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setGlobalNamedExpressions(newExpressions: Map<string, NamedExpression>): void {
    // Clear existing expressions without breaking the Map reference
    this.globalNamedExpressions.clear();
    
    // Repopulate with new expressions
    newExpressions.forEach((expression, name) => {
      this.globalNamedExpressions.set(name, expression);
    });
    
    this.eventEmitter?.emit("global-named-expressions-updated", this.globalNamedExpressions);
  }

  /**
   * Replace all sheet-scoped named expressions for a specific sheet (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setNamedExpressions(sheetName: string, newExpressions: Map<string, NamedExpression>): void {
    // Get or create the sheet's named expressions Map
    let sheetExpressions = this.scopedNamedExpressions.get(sheetName);
    if (!sheetExpressions) {
      sheetExpressions = new Map();
      this.scopedNamedExpressions.set(sheetName, sheetExpressions);
    }
    
    // Clear existing expressions without breaking the Map reference
    sheetExpressions.clear();
    
    // Repopulate with new expressions
    newExpressions.forEach((expression, name) => {
      sheetExpressions!.set(name, expression);
    });
    
    // Note: No specific event for scoped named expressions, but global event covers the change
    this.eventEmitter?.emit("global-named-expressions-updated", this.globalNamedExpressions);
  }
}
