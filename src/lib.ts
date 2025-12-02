export { FormulaEngine } from "./core/engine";
export * from "./core/types";
export * from "./core/utils";
export * from "./core/utils/color-utils";

// Schema exports
export { defineSchema } from "./core/schema/schema";
export type { CreateSchema, SchemaDeclaration, Schema, TableSchemaDefinition, CellSchemaDefinition, TableOrmSchema, CellOrmSchema } from "./core/schema/schema";
export { TableOrm } from "./core/schema/table-orm";
export { CellOrm } from "./core/schema/cell-orm";
export { SchemaValidationError } from "./core/managers/schema-manager";
export type { ValidationResult } from "./core/managers/schema-manager";

// Command Pattern exports
export { SchemaIntegrityError } from "./core/commands/command-executor";
export type { EngineAction, EngineCommand } from "./core/commands/types";
