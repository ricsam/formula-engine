/**
 * Schema Builder - Constructs the working schema from declarations
 *
 * This module is responsible for creating the actual working schema
 * from the schema declarations when attached to a FormulaEngine.
 */

import type { FormulaEngine } from "../engine";
import type { SchemaDeclaration, TableSchemaDefinition, CellSchemaDefinition, Schema } from "./schema";
import { TableOrm } from "./table-orm";
import { CellOrm } from "./cell-orm";
import type { SchemaManager } from "../managers/schema-manager";

/**
 * Build the working schema surface from declarations
 *
 * This creates TableOrm and CellOrm instances for each declared schema
 * and returns them directly.
 */
export function buildSchemaFromDeclaration(
  engine: FormulaEngine<any, any>,
  declaration: SchemaDeclaration,
  schemaManager: SchemaManager
): Schema {
  const schema: Schema = {};

  for (const [namespace, def] of Object.entries(declaration)) {
    if (def.type === "table") {
      schema[namespace] = buildTableSchema(engine, namespace, def, schemaManager);
    } else if (def.type === "cell") {
      schema[namespace] = buildCellSchema(engine, namespace, def, schemaManager);
    }
  }

  return schema;
}

/**
 * Build schema for a table schema - returns TableOrm instance directly
 */
function buildTableSchema(
  engine: FormulaEngine<any, any>,
  namespace: string,
  def: TableSchemaDefinition,
  schemaManager: SchemaManager
): TableOrm<any> {
  // Register the schema with the schema manager
  schemaManager.registerTableSchema(
    namespace,
    def.workbookName,
    def.tableName,
    def.headers
  );

  // Create and return the ORM instance directly
  return new TableOrm(
    engine,
    def.workbookName,
    def.tableName,
    def.headers,
    namespace
  );
}

/**
 * Build schema for a cell schema - returns CellOrm instance directly
 */
function buildCellSchema(
  engine: FormulaEngine<any, any>,
  namespace: string,
  def: CellSchemaDefinition,
  schemaManager: SchemaManager
): CellOrm<any> {
  // Register the schema with the schema manager
  schemaManager.registerCellSchema(namespace, def.cellAddress, def.parse);

  // Create and return the ORM instance directly
  return new CellOrm(engine, def.cellAddress, def.parse, namespace);
}
