/**
 * Named Expression Commands - Commands that modify named expressions
 *
 * These commands all require re-evaluation after execution.
 */

import type { NamedExpressionManager } from "../managers/named-expression-manager";
import type { WorkbookManager } from "../managers/workbook-manager";
import type { NamedExpression } from "../types";
import type {
  EngineCommand,
  EngineAction,
  MutationInvalidation,
} from "./types";
import { ActionTypes, emptyMutationInvalidation } from "./types";
import { getNamedExpressionResourceKey } from "../resource-keys";

/**
 * Dependencies needed for named expression commands.
 */
export interface NamedExpressionCommandDeps {
  namedExpressionManager: NamedExpressionManager;
  workbookManager: WorkbookManager;
  renameNamedExpressionInFormula: (
    formula: string,
    oldName: string,
    newName: string
  ) => string;
}

/**
 * Helper to convert opts to getNamedExpression format
 */
function optsToScope(opts: {
  expressionName: string;
  sheetName?: string;
  workbookName?: string;
}): {
  name: string;
  scope:
    | { type: "global" }
    | { type: "workbook"; workbookName: string }
    | { type: "sheet"; workbookName: string; sheetName: string };
} {
  if (opts.sheetName && opts.workbookName) {
    return {
      name: opts.expressionName,
      scope: {
        type: "sheet",
        workbookName: opts.workbookName,
        sheetName: opts.sheetName,
      },
    };
  } else if (opts.workbookName) {
    return {
      name: opts.expressionName,
      scope: { type: "workbook", workbookName: opts.workbookName },
    };
  } else {
    return {
      name: opts.expressionName,
      scope: { type: "global" },
    };
  }
}

function getNamedExpressionScopeResourceKeys(
  expressions: Iterable<string>,
  opts: {
    workbookName?: string;
    sheetName?: string;
  }
): string[] {
  return Array.from(
    new Set(
      Array.from(expressions, (expressionName) =>
        getNamedExpressionResourceKey({
          expressionName,
          workbookName: opts.workbookName,
          sheetName: opts.sheetName,
        })
      )
    )
  );
}

/**
 * Command to add a named expression.
 */
export class AddNamedExpressionCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: NamedExpressionCommandDeps,
    private opts: {
      expression: string;
      expressionName: string;
      sheetName?: string;
      workbookName?: string;
    }
  ) {}

  execute(): void {
    this.deps.namedExpressionManager.addNamedExpression(this.opts);
    const resourceKey = getNamedExpressionResourceKey(this.opts);
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey],
    };
  }

  undo(): void {
    this.deps.namedExpressionManager.removeNamedExpression({
      expressionName: this.opts.expressionName,
      sheetName: this.opts.sheetName,
      workbookName: this.opts.workbookName,
    });
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.ADD_NAMED_EXPRESSION,
      payload: this.opts,
    };
  }
}

/**
 * Command to remove a named expression.
 */
export class RemoveNamedExpressionCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private removedExpression: NamedExpression | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: NamedExpressionCommandDeps,
    private opts: {
      expressionName: string;
      sheetName?: string;
      workbookName?: string;
    }
  ) {}

  execute(): void {
    // Capture expression before removal
    this.removedExpression = this.deps.namedExpressionManager.getNamedExpression(
      optsToScope(this.opts)
    );

    this.deps.namedExpressionManager.removeNamedExpression(this.opts);
    const resourceKey = getNamedExpressionResourceKey(this.opts);
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey],
    };
  }

  undo(): void {
    if (!this.removedExpression) return;

    this.deps.namedExpressionManager.addNamedExpression({
      expressionName: this.removedExpression.name,
      expression: this.removedExpression.expression,
      sheetName: this.opts.sheetName,
      workbookName: this.opts.workbookName,
    });
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.REMOVE_NAMED_EXPRESSION,
      payload: this.opts,
    };
  }
}

/**
 * Command to update a named expression.
 */
export class UpdateNamedExpressionCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousExpression: string | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: NamedExpressionCommandDeps,
    private opts: {
      expression: string;
      expressionName: string;
      sheetName?: string;
      workbookName?: string;
    }
  ) {}

  execute(): void {
    // Capture previous expression
    const existing = this.deps.namedExpressionManager.getNamedExpression(
      optsToScope(this.opts)
    );
    this.previousExpression = existing?.expression;

    this.deps.namedExpressionManager.updateNamedExpression(this.opts);
    const resourceKey = getNamedExpressionResourceKey(this.opts);
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey],
    };
  }

  undo(): void {
    if (this.previousExpression === undefined) return;

    this.deps.namedExpressionManager.updateNamedExpression({
      expressionName: this.opts.expressionName,
      expression: this.previousExpression,
      sheetName: this.opts.sheetName,
      workbookName: this.opts.workbookName,
    });
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.UPDATE_NAMED_EXPRESSION,
      payload: this.opts,
    };
  }
}

/**
 * Command to rename a named expression.
 */
export class RenameNamedExpressionCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: NamedExpressionCommandDeps,
    private opts: {
      expressionName: string;
      sheetName?: string;
      workbookName?: string;
      newName: string;
    }
  ) {}

  execute(): void {
    this.deps.namedExpressionManager.renameNamedExpression(this.opts);

    // Update formulas in sheet cells
    const changedCells = this.deps.workbookManager.updateAllFormulas((formula) =>
      this.deps.renameNamedExpressionInFormula(
        formula,
        this.opts.expressionName,
        this.opts.newName
      )
    );

    // Update named expressions
    const changedNamedExpressions =
      this.deps.namedExpressionManager.updateAllNamedExpressions((formula) =>
      this.deps.renameNamedExpressionInFormula(
        formula,
        this.opts.expressionName,
        this.opts.newName
      )
    );

    this.executeFootprint = {
      touchedCells: changedCells.map((address) => ({
        address,
        beforeKind: "formula" as const,
        afterKind: "formula" as const,
      })),
      resourceKeys: [
        getNamedExpressionResourceKey({
          expressionName: this.opts.expressionName,
          workbookName: this.opts.workbookName,
          sheetName: this.opts.sheetName,
        }),
        getNamedExpressionResourceKey({
          expressionName: this.opts.newName,
          workbookName: this.opts.workbookName,
          sheetName: this.opts.sheetName,
        }),
        ...changedNamedExpressions,
      ],
    };
  }

  undo(): void {
    // Rename back
    this.deps.namedExpressionManager.renameNamedExpression({
      expressionName: this.opts.newName,
      sheetName: this.opts.sheetName,
      workbookName: this.opts.workbookName,
      newName: this.opts.expressionName,
    });

    // Update formulas back
    const changedCells = this.deps.workbookManager.updateAllFormulas((formula) =>
      this.deps.renameNamedExpressionInFormula(
        formula,
        this.opts.newName,
        this.opts.expressionName
      )
    );

    // Update named expressions back
    const changedNamedExpressions =
      this.deps.namedExpressionManager.updateAllNamedExpressions((formula) =>
      this.deps.renameNamedExpressionInFormula(
        formula,
        this.opts.newName,
        this.opts.expressionName
      )
    );

    this.undoFootprint = {
      touchedCells: changedCells.map((address) => ({
        address,
        beforeKind: "formula" as const,
        afterKind: "formula" as const,
      })),
      resourceKeys: [
        getNamedExpressionResourceKey({
          expressionName: this.opts.expressionName,
          workbookName: this.opts.workbookName,
          sheetName: this.opts.sheetName,
        }),
        getNamedExpressionResourceKey({
          expressionName: this.opts.newName,
          workbookName: this.opts.workbookName,
          sheetName: this.opts.sheetName,
        }),
        ...changedNamedExpressions,
      ],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.RENAME_NAMED_EXPRESSION,
      payload: this.opts,
    };
  }
}

/**
 * Options for setNamedExpressions command.
 */
type SetNamedExpressionsOpts = (
  | { type: "global" }
  | { type: "sheet"; sheetName: string; workbookName: string }
  | { type: "workbook"; workbookName: string }
) & {
  expressions: Map<string, NamedExpression>;
};

/**
 * Command to set named expressions (replace all at a scope).
 */
export class SetNamedExpressionsCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousExpressions: Map<string, NamedExpression> | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: NamedExpressionCommandDeps,
    private opts: SetNamedExpressionsOpts
  ) {}

  execute(): void {
    // Capture previous expressions at this scope
    const allExpressions = this.deps.namedExpressionManager.getNamedExpressions();

    if (this.opts.type === "global") {
      this.previousExpressions = new Map(allExpressions.globalExpressions);
    } else if (this.opts.type === "workbook") {
      this.previousExpressions = new Map(
        allExpressions.workbookExpressions.get(this.opts.workbookName) || []
      );
    } else if (this.opts.type === "sheet") {
      const sheetExpressions = allExpressions.sheetExpressions
        .get(this.opts.workbookName)
        ?.get(this.opts.sheetName);
      this.previousExpressions = new Map(sheetExpressions || []);
    }

    this.deps.namedExpressionManager.setNamedExpressions(this.opts);
    const scope =
      this.opts.type === "global"
        ? {}
        : this.opts.type === "workbook"
        ? { workbookName: this.opts.workbookName }
        : {
            workbookName: this.opts.workbookName,
            sheetName: this.opts.sheetName,
          };
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [
        ...getNamedExpressionScopeResourceKeys(
          this.previousExpressions?.keys() ?? [],
          scope
        ),
        ...getNamedExpressionScopeResourceKeys(this.opts.expressions.keys(), scope),
      ],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: this.executeFootprint.resourceKeys,
    };
  }

  undo(): void {
    if (!this.previousExpressions) return;

    // Restore previous expressions
    if (this.opts.type === "global") {
      this.deps.namedExpressionManager.setNamedExpressions({
        type: "global",
        expressions: this.previousExpressions,
      });
    } else if (this.opts.type === "workbook") {
      this.deps.namedExpressionManager.setNamedExpressions({
        type: "workbook",
        workbookName: this.opts.workbookName,
        expressions: this.previousExpressions,
      });
    } else if (this.opts.type === "sheet") {
      this.deps.namedExpressionManager.setNamedExpressions({
        type: "sheet",
        workbookName: this.opts.workbookName,
        sheetName: this.opts.sheetName,
        expressions: this.previousExpressions,
      });
    }
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.SET_NAMED_EXPRESSIONS,
      payload: {
        ...this.opts,
        expressions: Array.from(this.opts.expressions.entries()),
      },
    };
  }
}
