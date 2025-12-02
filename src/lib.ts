export { FormulaEngine } from "./core/engine";
export * from "./core/types";
export * from "./core/utils";
export * from "./core/utils/color-utils";

// API Schema exports
export { defineApi, createApi } from "./core/api/api";
export type { CreateApi, Declaration, Api, TableApi, CellApi } from "./core/api/api";
export { TableOrm } from "./core/api/table-orm";
export { CellOrm } from "./core/api/cell-orm";
export { SchemaValidationError } from "./core/managers/api-schema-manager";
export type { ValidationResult } from "./core/managers/api-schema-manager";

// Command Pattern exports
export { SchemaIntegrityError } from "./core/commands/command-executor";
export type { EngineAction, EngineCommand } from "./core/commands/types";
