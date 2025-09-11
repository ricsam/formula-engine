import type { NamedExpression } from "../types";
import { renameNamedExpressionInFormula } from "../named-expression-renamer";
import type { EventManager } from "./event-manager";

export class NamedExpressionManager {
  private sheetExpressions: Map<
    string,
    Map<string, Map<string, NamedExpression>>
  > = new Map();
  private workbookExpressions: Map<string, Map<string, NamedExpression>> =
    new Map();
  private globalExpressions: Map<string, NamedExpression> = new Map();
  private eventEmitter: EventManager;

  constructor(eventEmitter: EventManager) {
    this.eventEmitter = eventEmitter;
  }

  getSheetNamedExpressions(): Map<
    string,
    Map<string, Map<string, NamedExpression>>
  > {
    return this.sheetExpressions;
  }

  getWorkbookNamedExpressions(): Map<string, Map<string, NamedExpression>> {
    return this.workbookExpressions;
  }

  getGlobalNamedExpressions(): Map<string, NamedExpression> {
    return this.globalExpressions;
  }

  addNamedExpression({
    expression,
    expressionName,
    sheetName,
    workbookName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    const namedExpression: NamedExpression = {
      name: expressionName,
      expression,
      sheetName,
      workbookName,
    };

    if (sheetName && !workbookName) {
      throw new Error("Missing workbookName");
    }

    if (sheetName && workbookName) {
      let wbLevel = this.sheetExpressions.get(workbookName);
      if (!wbLevel) {
        wbLevel = new Map();
        this.sheetExpressions.set(sheetName, wbLevel);
      }

      let sheetLevel = wbLevel.get(sheetName);
      if (!sheetLevel) {
        sheetLevel = new Map();
        wbLevel.set(sheetName, sheetLevel);
      }

      sheetLevel.set(expressionName, namedExpression);
    } else if (workbookName) {
      let workbookNamedExpressions = this.workbookExpressions.get(workbookName);
      if (!workbookNamedExpressions) {
        workbookNamedExpressions = new Map();
        this.workbookExpressions.set(workbookName, workbookNamedExpressions);
      }

      workbookNamedExpressions.set(expressionName, namedExpression);
    } else {
      this.globalExpressions.set(expressionName, namedExpression);
      this.eventEmitter.emit(
        "global-named-expressions-updated",
        this.globalExpressions
      );
    }
  }

  removeNamedExpression({
    expressionName,
    sheetName,
    workbookName,
  }: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): boolean {
    let found = false;

    if (sheetName && !workbookName) {
      throw new Error("Missing workbookName");
    }

    if (sheetName && workbookName) {
      const wbLevel = this.sheetExpressions.get(workbookName);
      if (wbLevel) {
        const sheetLevel = wbLevel.get(sheetName);
        if (sheetLevel) {
          found = sheetLevel.delete(expressionName);
        }
      }
    } else if (workbookName) {
      const workbookNamedExpressions =
        this.workbookExpressions.get(workbookName);
      if (workbookNamedExpressions) {
        found = workbookNamedExpressions.delete(expressionName);
      }
    } else {
      found = this.globalExpressions.delete(expressionName);
      if (found) {
        this.eventEmitter.emit(
          "global-named-expressions-updated",
          this.globalExpressions
        );
      }
    }

    return found;
  }

  updateNamedExpression({
    expression,
    expressionName,
    sheetName,
    workbookName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    // Check if the named expression exists
    let exists = false;

    if (sheetName && !workbookName) {
      throw new Error("Missing workbookName");
    }

    if (sheetName && workbookName) {
      const wbLevel = this.sheetExpressions.get(workbookName);
      if (wbLevel) {
        const sheetLevel = wbLevel.get(sheetName);
        if (sheetLevel) {
          exists = sheetLevel.has(expressionName);
        }
      }
    } else if (workbookName) {
      const workbookNamedExpressions =
        this.workbookExpressions.get(workbookName);
      if (workbookNamedExpressions) {
        exists = workbookNamedExpressions.has(expressionName);
      }
    } else {
      exists = this.globalExpressions.has(expressionName);
    }

    if (!exists) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Update is the same as add for existing expressions
    this.addNamedExpression({ expression, expressionName, sheetName });
  }

  renameNamedExpression({
    expressionName,
    sheetName,
    workbookName,
    newName,
  }: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
    newName: string;
  }): boolean {
    // Check if the named expression exists
    let targetMap: Map<string, NamedExpression> | undefined;
    if (sheetName && !workbookName) {
      throw new Error("Missing workbookName");
    }

    let isGlobal = false;

    if (sheetName && workbookName) {
      const wbLevel = this.sheetExpressions.get(workbookName);
      if (wbLevel) {
        const sheetLevel = wbLevel.get(sheetName);
        if (sheetLevel) {
          targetMap = sheetLevel;
        }
      }
    } else if (workbookName) {
      targetMap = this.workbookExpressions.get(workbookName);
    } else {
      targetMap = this.globalExpressions;
      isGlobal = true;
    }

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

    if (isGlobal) {
      this.eventEmitter.emit(
        "global-named-expressions-updated",
        this.globalExpressions
      );
    }

    return true;
  }

  updateAllNamedExpressions(updateCallback: (formula: string) => string): void {
    const update = (map: Map<string, NamedExpression>) => {
      map.forEach((namedExpr, name) => {
        // Don't update the expression we're renaming
        const updatedExpression = updateCallback(namedExpr.expression);

        if (updatedExpression !== namedExpr.expression) {
          map.set(name, {
            ...namedExpr,
            expression: updatedExpression,
          });
        }
      });
    };

    update(this.globalExpressions);

    this.workbookExpressions.forEach((workbookLevel) => {
      update(workbookLevel);
    });

    this.sheetExpressions.forEach((wbLevel) => {
      wbLevel.forEach((sheetLevel) => {
        update(sheetLevel);
      });
    });
  }

  getSheetExpressionsSerialized({
    sheetName,
    workbookName,
  }: {
    sheetName: string;
    workbookName: string;
  }): Map<string, NamedExpression> {
    const wbLevel = this.sheetExpressions.get(workbookName);
    if (wbLevel) {
      const sheetLevel = wbLevel.get(sheetName);
      if (sheetLevel) {
        return sheetLevel;
      }
    }
    throw new Error("Named expressions not found");
  }

  getGlobalNamedExpressionsSerialized(): Map<string, NamedExpression> {
    return this.globalExpressions;
  }

  removeSheetExpressions(options: {
    sheetName: string;
    workbookName: string;
  }): void {
    const wbLevel = this.sheetExpressions.get(options.workbookName);
    if (wbLevel) {
      const sheetLevel = wbLevel.get(options.sheetName);
      if (sheetLevel) {
        sheetLevel.delete(options.sheetName);
      }
    }
    throw new Error("Named expressions not found");
  }

  renameSheetExpressions(options: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    const wbLevel = this.sheetExpressions.get(options.workbookName);
    if (wbLevel) {
      const sheetLevel = wbLevel.get(options.sheetName);
      if (sheetLevel) {
        sheetLevel.set(
          options.newSheetName,
          sheetLevel.get(options.sheetName)!
        );
        sheetLevel.delete(options.sheetName);
      }
    }
    throw new Error("Named expressions not found");
  }

  /**
   * Replace all global named expressions (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setGlobalNamedExpressions(
    newExpressions: Map<string, NamedExpression>
  ): void {
    // Clear existing expressions without breaking the Map reference
    this.globalExpressions.clear();

    // Repopulate with new expressions
    newExpressions.forEach((expression, name) => {
      this.globalExpressions.set(name, expression);
    });

    this.eventEmitter.emit(
      "global-named-expressions-updated",
      this.globalExpressions
    );
  }

  /**
   * Replace all sheet-scoped named expressions for a specific sheet (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setNamedExpressions(options: {
    sheetName: string;
    workbookName: string;
    newExpressions: Map<string, NamedExpression>;
  }): void {
    // Get or create the sheet's named expressions Map
    const wbLevel = this.sheetExpressions.get(options.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    let sheetExpressions = wbLevel.get(options.sheetName);
    if (!sheetExpressions) {
      sheetExpressions = new Map();
      wbLevel.set(options.sheetName, sheetExpressions);
    }

    // Clear existing expressions without breaking the Map reference
    sheetExpressions.clear();

    // Repopulate with new expressions
    options.newExpressions.forEach((expression, name) => {
      sheetExpressions!.set(name, expression);
    });

    // Note: No specific event for scoped named expressions, but global event covers the change
    this.eventEmitter.emit(
      "global-named-expressions-updated",
      this.globalExpressions
    );
  }

  getNamedExpression({
    sheetName,
    workbookName,
    name,
  }: {
    sheetName: string;
    workbookName: string;
    name: string;
  }): NamedExpression | undefined {
    if (sheetName && !workbookName) {
      throw new Error("Missing workbookName");
    }
    if (sheetName && workbookName) {
      const wbLevel = this.sheetExpressions.get(workbookName);
      if (wbLevel) {
        const sheetLevel = wbLevel.get(sheetName);
        if (sheetLevel) {
          return sheetLevel.get(name);
        }
      }
    } else if (workbookName) {
      const workbookLevel = this.workbookExpressions.get(workbookName);
      if (workbookLevel) {
        return workbookLevel.get(name);
      }
    } else {
      return this.globalExpressions.get(name);
    }
    return undefined;
  }
}
