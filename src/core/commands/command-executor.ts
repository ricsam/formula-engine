/**
 * CommandExecutor - Executes commands with undo/redo support and schema validation
 *
 * Handles:
 * - Command execution with optional re-evaluation
 * - Schema validation after re-evaluation
 * - Automatic rollback on validation failure
 * - Undo/redo stacks
 * - Action serialization for persistence
 */

import type { SchemaDeclaration } from "../schema/schema";
import type { FormulaEngine } from "../engine";
import type { EvaluationManager } from "../managers/evaluation-manager";
import type { EventManager } from "../managers/event-manager";
import type { CellAddress, SerializedCellValue } from "../types";
import type {
  EngineCommand,
  EngineAction,
  ExecuteOptions,
  SchemaValidationResult,
  SchemaValidationErrorInfo,
} from "./types";
import { emptyMutationInvalidation } from "./types";

/**
 * Error thrown when schema integrity validation fails.
 * Contains all validation errors that occurred.
 */
export class SchemaIntegrityError extends Error {
  constructor(public errors: SchemaValidationErrorInfo[]) {
    const messages = errors.map((e) => e.message).join("; ");
    super(`Schema integrity violation: ${messages}`);
    this.name = "SchemaIntegrityError";
  }
}

/**
 * CommandExecutor manages command execution with validation and undo/redo support.
 */
export class CommandExecutor {
  private undoStack: EngineCommand[] = [];
  private redoStack: EngineCommand[] = [];
  private actionLog: EngineAction[] = [];

  constructor(
    private evaluationManager: EvaluationManager,
    private eventManager: EventManager,
    private validateAllSchemas: () => SchemaValidationResult
  ) {}

  /**
   * Execute a command with optional schema validation.
   *
   * @param command - The command to execute
   * @param options - Execution options
   * @throws SchemaIntegrityError if validation fails
   */
  execute(command: EngineCommand, options: ExecuteOptions = {}): void {
    const {
      validate = false,
      skipUndoStack = false,
      skipEmitUpdate = false,
    } = options;

    // Execute the command
    command.execute();

    // Re-evaluate if needed
    if (command.requiresReevaluation) {
      this.evaluationManager.invalidateFromMutation(
        command.getInvalidationFootprint?.("execute") ??
          emptyMutationInvalidation()
      );

      // Validate schemas if requested
      if (validate) {
        const validation = this.validateAllSchemas();

        if (!validation.valid) {
          // Rollback: undo the command and re-evaluate
          command.undo();
          this.evaluationManager.invalidateFromMutation(
            command.getInvalidationFootprint?.("undo") ??
              emptyMutationInvalidation()
          );
          throw new SchemaIntegrityError(validation.errors);
        }
      }
    }

    // Add to undo stack
    if (!skipUndoStack) {
      this.undoStack.push(command);
      // Clear redo stack on new action
      this.redoStack = [];
    }

    // Log the action
    const action = command.toAction();
    action.timestamp = Date.now();
    this.actionLog.push(action);

    // Emit update event
    if (!skipEmitUpdate) {
      this.eventManager.emitUpdate();
    }
  }

  /**
   * Undo the last command.
   *
   * @returns true if undo was performed, false if nothing to undo
   */
  undo(): boolean {
    const command = this.undoStack.pop();
    if (!command) {
      return false;
    }

    command.undo();

    if (command.requiresReevaluation) {
      this.evaluationManager.invalidateFromMutation(
        command.getInvalidationFootprint?.("undo") ??
          emptyMutationInvalidation()
      );
    }

    this.redoStack.push(command);
    this.eventManager.emitUpdate();

    return true;
  }

  /**
   * Redo the last undone command.
   *
   * @returns true if redo was performed, false if nothing to redo
   */
  redo(): boolean {
    const command = this.redoStack.pop();
    if (!command) {
      return false;
    }

    command.execute();

    if (command.requiresReevaluation) {
      this.evaluationManager.invalidateFromMutation(
        command.getInvalidationFootprint?.("execute") ??
          emptyMutationInvalidation()
      );
    }

    this.undoStack.push(command);
    this.eventManager.emitUpdate();

    return true;
  }

  /**
   * Check if undo is available.
   */
  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  /**
   * Check if redo is available.
   */
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Get the action log for persistence/collaboration.
   */
  getActionLog(): EngineAction[] {
    return [...this.actionLog];
  }

  /**
   * Clear the action log.
   */
  clearActionLog(): void {
    this.actionLog = [];
  }

  /**
   * Clear undo/redo stacks.
   */
  clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  /**
   * Get the number of commands in the undo stack.
   */
  getUndoStackSize(): number {
    return this.undoStack.length;
  }

  /**
   * Get the number of commands in the redo stack.
   */
  getRedoStackSize(): number {
    return this.redoStack.length;
  }
}
