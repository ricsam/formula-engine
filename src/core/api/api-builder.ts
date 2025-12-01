/**
 * API Builder - Constructs the working API from declarations
 *
 * This module is responsible for creating the actual working API
 * from the schema declarations when attached to a FormulaEngine.
 */

import type { FormulaEngine } from "../engine";
import type { Declaration, TableApi, CellApi, Api } from "./api";
import { TableOrm } from "./table-orm";
import { CellOrm } from "./cell-orm";
import type { ApiSchemaManager } from "../managers/api-schema-manager";

/**
 * Build the working API surface from declarations
 *
 * This creates TableOrm and CellOrm instances for each declared schema
 * and binds the custom methods to them.
 */
export function buildApiFromDeclaration(
  engine: FormulaEngine<any, any>,
  declaration: Declaration,
  schemaManager: ApiSchemaManager
): Api {
  const api: Api = {};

  for (const [namespace, def] of Object.entries(declaration)) {
    if (def.type === "table") {
      api[namespace] = buildTableApi(engine, namespace, def, schemaManager);
    } else if (def.type === "cell") {
      api[namespace] = buildCellApi(engine, namespace, def, schemaManager);
    }
  }

  return api;
}

/**
 * Build API methods for a table schema
 */
function buildTableApi(
  engine: FormulaEngine<any, any>,
  namespace: string,
  def: TableApi,
  schemaManager: ApiSchemaManager
): Record<string, (...args: any[]) => any> {
  // Register the schema with the schema manager
  schemaManager.registerTableSchema(
    namespace,
    def.workbookName,
    def.tableName,
    def.headers
  );

  // Create the ORM instance
  const orm = new TableOrm(
    engine,
    def.workbookName,
    def.tableName,
    def.headers,
    namespace
  );

  // Bind all custom methods to the ORM instance
  const boundMethods: Record<string, (...args: any[]) => any> = {};

  for (const [methodName, methodFn] of Object.entries(def.methods)) {
    boundMethods[methodName] = methodFn.bind(orm);
  }

  return boundMethods;
}

/**
 * Build API methods for a cell schema
 */
function buildCellApi(
  engine: FormulaEngine<any, any>,
  namespace: string,
  def: CellApi,
  schemaManager: ApiSchemaManager
): Record<string, (...args: any[]) => any> {
  // Register the schema with the schema manager
  schemaManager.registerCellSchema(namespace, def.cellAddress, def.parse);

  // Create the ORM instance
  const orm = new CellOrm(engine, def.cellAddress, def.parse, namespace);

  // Bind all custom methods to the ORM instance
  const boundMethods: Record<string, (...args: any[]) => any> = {};

  for (const [methodName, methodFn] of Object.entries(def.methods)) {
    boundMethods[methodName] = methodFn.bind(orm);
  }

  return boundMethods;
}

