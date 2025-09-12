import type { EvaluationContext, NamedExpression } from "../types";
import { renameNamedExpressionInFormula } from "../named-expression-renamer";
import type { EventManager } from "./event-manager";
import type { NamedExpressionNode } from "src/parser/ast";

export class NamedExpressionManager {
  sheetExpressions: Map<string, Map<string, Map<string, NamedExpression>>> =
    new Map();
  workbookExpressions: Map<string, Map<string, NamedExpression>> = new Map();
  globalExpressions: Map<string, NamedExpression> = new Map();

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
    this.addNamedExpression({
      expression,
      expressionName,
      sheetName,
      workbookName,
    });
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

  /**
   * Replace all named expressions
   */
  setNamedExpressions(
    opts: (
      | {
          type: "global";
        }
      | {
          type: "sheet";
          sheetName: string;
          workbookName: string;
        }
      | {
          type: "workbook";
          workbookName: string;
        }
    ) & {
      expressions: Map<string, NamedExpression>;
    }
  ) {
    let map: Map<string, NamedExpression> | undefined;

    if (opts.type === "sheet") {
      map = this.sheetExpressions.get(opts.workbookName)?.get(opts.sheetName);
    } else if (opts.type === "workbook") {
      map = this.workbookExpressions.get(opts.workbookName);
    } else {
      map = this.globalExpressions;
    }

    if (!map) {
      throw new Error("Invalid options: " + JSON.stringify(opts));
    }

    map.clear();

    opts.expressions.forEach((expression, name) => {
      map.set(name, expression);
    });
  }

  getNamedExpression(depNode: {
    name: string;
    scope:
      | {
          type: "global";
        }
      | {
          type: "workbook";
          workbookName: string;
        }
      | {
          type: "sheet";
          workbookName: string;
          sheetName: string;
        };
  }): NamedExpression | undefined {
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
      NamedExpressionNode,
      "name" | "sheetName" | "workbookName"
    >,
    context: EvaluationContext
  ): string | undefined {
    // scenario 1: no sheetName nor workbookName
    if (!namedExpression.sheetName && !namedExpression.workbookName) {
      // step 1, check if there is a named expression in the sheet scope
      const expression = this.sheetExpressions
        .get(context.currentCell.workbookName)
        ?.get(context.currentCell.sheetName)
        ?.get(namedExpression.name);
      if (expression) {
        return expression.expression;
      } else {
        // step 2, check if there is a named expression in the workbook scope
        const expression = this.workbookExpressions
          .get(context.currentCell.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return expression.expression;
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return expression.expression;
          }
        }
      }
    }

    // scenario 2: we only have a workbookName - a bit weird, but could happen
    if (namedExpression.workbookName && !namedExpression.sheetName) {
      // special case: if workbook is the current workbook, we should just resolve the named expression according to scenario 1
      if (namedExpression.workbookName === context.currentCell.workbookName) {
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
        return expression.expression;
      } else {
        // step 2, check if there is a named expression in the global scope
        const expression = this.globalExpressions.get(namedExpression.name);
        if (expression) {
          return expression.expression;
        }
      }
    }

    // scenario 3: we only have a sheetName
    if (namedExpression.sheetName && !namedExpression.workbookName) {
      const expression = this.sheetExpressions
        .get(context.currentCell.workbookName)
        ?.get(namedExpression.sheetName)
        ?.get(namedExpression.name);
      if (expression) {
        // step 1, check if there is a named expression in the current workbook against the sheet name
        return expression.expression;
      } else {
        // step 2, check if there is a named expression in the current workbook has a workbook scoped named expression
        const expression = this.workbookExpressions
          .get(context.currentCell.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return expression.expression;
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return expression.expression;
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
        return expression.expression;
      } else {
        // step 2, check if there is a named expression in the workbook scope
        const expression = this.workbookExpressions
          .get(namedExpression.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return expression.expression;
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return expression.expression;
          }
        }
      }
    }
  }

  getNamedExpressions() {
    return {
      sheetExpressions: this.sheetExpressions,
      workbookExpressions: this.workbookExpressions,
      globalExpressions: this.globalExpressions,
    };
  }

  resetNamedExpressions(
    namedExpressions: ReturnType<typeof this.getNamedExpressions>
  ) {
    this.setNamedExpressions({
      type: "global",
      expressions: namedExpressions.globalExpressions,
    });

    namedExpressions.workbookExpressions.forEach(
      (workbookExpressions, workbookName) => {
        this.setNamedExpressions({
          type: "workbook",
          expressions: workbookExpressions,
          workbookName,
        });
      }
    );

    namedExpressions.sheetExpressions.forEach(
      (sheetExpressions, workbookName) => {
        sheetExpressions.forEach((sheetExpression, sheetName) => {
          this.setNamedExpressions({
            type: "sheet",
            expressions: sheetExpression,
            sheetName,
            workbookName,
          });
        });
      }
    );
  }

  /**
   * When adding a sheet, we need to initialize the new maps
   */
  addSheet(opts: { workbookName: string; sheetName: string }) {
    const wbLevel = this.sheetExpressions.get(opts.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    const sheetLevel = wbLevel.get(opts.sheetName);
    if (sheetLevel) {
      throw new Error("Sheet already exists");
    }
    wbLevel.set(opts.sheetName, new Map());
  }

  /**
   * When adding a workbook, we need to initialize the new maps
   */
  addWorkbook(workbookName: string) {
    this.sheetExpressions.set(workbookName, new Map());
    this.workbookExpressions.set(workbookName, new Map());
  }

  /**
   * When removing a workbook, we need to remove the workbook from the sheet level
   */
  removeWorkbook(workbookName: string) {
    this.sheetExpressions.delete(workbookName);
    this.workbookExpressions.delete(workbookName);
  }

  /**
   * When removing a sheet, we need to remove the sheet from the workbook level
   */
  removeSheet(opts: { workbookName: string; sheetName: string }) {
    const wbLevel = this.sheetExpressions.get(opts.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    wbLevel.delete(opts.sheetName);
  }

  /**
   * Rename a sheet's named expressions, mainly used when renaming a sheet
   */
  renameSheet(options: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    const wbLevel = this.sheetExpressions.get(options.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    const sheetLevel = wbLevel.get(options.sheetName);
    if (!sheetLevel) {
      throw new Error("Sheet not found");
    }
    wbLevel.set(options.newSheetName, sheetLevel);
    wbLevel.delete(options.sheetName);
  }
}
