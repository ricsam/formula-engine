import type {
  EvaluationContext,
  NamedExpression,
  NamedExpressionDependencyNode,
} from "../types";
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
        this.sheetExpressions.set(workbookName, wbLevel);
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
    this.addNamedExpression({ expression, expressionName, sheetName, workbookName });
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
    return new Map();
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
      wbLevel.delete(options.sheetName);
    }
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
        wbLevel.set(options.newSheetName, sheetLevel);
        wbLevel.delete(options.sheetName);
      }
    }
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

  getNamedExpression(
    depNode: NamedExpressionDependencyNode
  ): NamedExpression | undefined {
    if (depNode.scope.type === "global") {
      return this.globalExpressions.get(depNode.name);
    }
    if (depNode.scope.type === "workbook") {
      return this.workbookExpressions
        .get(depNode.scope.workbookName)
        ?.get(depNode.name);
    }
    if (depNode.scope.type === "sheet") {
      return this.sheetExpressions
        .get(depNode.scope.workbookName)
        ?.get(depNode.scope.sheetName)
        ?.get(depNode.name);
    }
    return undefined;
  }

  resolveNamedExpression(
    namedExpression: Pick<
      NamedExpression,
      "name" | "sheetName" | "workbookName"
    >,
    context: EvaluationContext
  ): NamedExpressionDependencyNode | undefined {
    // scenario 1: no sheetName nor workbookName
    if (!namedExpression.sheetName && !namedExpression.workbookName) {
      // step 1, check if there is a named expression in the sheet scope
      const expression = this.sheetExpressions
        .get(context.currentWorkbook)
        ?.get(context.currentSheet)
        ?.get(namedExpression.name);
      if (expression) {
        return {
          type: "named-expression",
          name: expression.name,
          scope: {
            type: "sheet",
            sheetName: context.currentSheet,
            workbookName: context.currentWorkbook,
          },
        };
      } else {
        // step 2, check if there is a named expression in the workbook scope
        const expression = this.workbookExpressions
          .get(context.currentWorkbook)
          ?.get(namedExpression.name);
        if (expression) {
          return {
            type: "named-expression",
            name: expression.name,
            scope: { type: "workbook", workbookName: context.currentWorkbook },
          };
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return {
              type: "named-expression",
              name: expression.name,
              scope: { type: "global" },
            };
          }
        }
      }
    }

    // scenario 2: we only have a workbookName - a bit weird, but could happen
    if (namedExpression.workbookName && !namedExpression.sheetName) {
      // special case: if workbook is the current workbook, we should just resolve the named expression according to scenario 1
      if (namedExpression.workbookName === context.currentWorkbook) {
        return this.resolveNamedExpression(
          {
            name: namedExpression.name,
          },
          context
        );
      }

      const expression = this.workbookExpressions
        .get(namedExpression.workbookName)
        ?.get(namedExpression.name);
      if (expression) {
        // step 1, check if there is a named expression in the workbook scope
        return {
          type: "named-expression",
          name: expression.name,
          scope: {
            type: "workbook",
            workbookName: namedExpression.workbookName,
          },
        };
      } else {
        // step 2, check if there is a named expression in the global scope
        const expression = this.globalExpressions.get(namedExpression.name);
        if (expression) {
          return {
            type: "named-expression",
            name: expression.name,
            scope: { type: "global" },
          };
        }
      }
    }

    // scenario 3: we only have a sheetName
    if (namedExpression.sheetName && !namedExpression.workbookName) {
      const expression = this.sheetExpressions
        .get(context.currentWorkbook)
        ?.get(namedExpression.sheetName)
        ?.get(namedExpression.name);
      if (expression) {
        // step 1, check if there is a named expression in the current workbook against the sheet name
        return {
          type: "named-expression",
          name: expression.name,
          scope: {
            type: "sheet",
            sheetName: namedExpression.sheetName,
            workbookName: context.currentWorkbook,
          },
        };
      } else {
        // step 2, check if there is a named expression in the current workbook has a workbook scoped named expression
        const expression = this.workbookExpressions
          .get(context.currentWorkbook)
          ?.get(namedExpression.name);
        if (expression) {
          return {
            type: "named-expression",
            name: expression.name,
            scope: { type: "workbook", workbookName: context.currentWorkbook },
          };
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return {
              type: "named-expression",
              name: expression.name,
              scope: { type: "global" },
            };
          }
        }
      }
    }

    // scenario 4: we have both sheetName and workbookName
    if (namedExpression.sheetName && namedExpression.workbookName) {
      const expression = this.sheetExpressions
        .get(namedExpression.workbookName)
        ?.get(namedExpression.sheetName)
        ?.get(namedExpression.name);
      if (expression) {
        // step 1, check if there is a named expression the the sheet scope
        return {
          type: "named-expression",
          name: expression.name,
          scope: {
            type: "sheet",
            sheetName: namedExpression.sheetName,
            workbookName: namedExpression.workbookName,
          },
        };
      } else {
        // step 2, check if there is a named expression in the workbook scope
        const expression = this.workbookExpressions
          .get(namedExpression.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return {
            type: "named-expression",
            name: expression.name,
            scope: {
              type: "workbook",
              workbookName: namedExpression.workbookName,
            },
          };
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return {
              type: "named-expression",
              name: expression.name,
              scope: { type: "global" },
            };
          }
        }
      }
    }
  }
}
