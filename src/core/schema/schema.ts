import type { CellAddress } from "../types";

/**
 * Define a schema for the FormulaEngine.
 *
 * The returned object contains:
 * - `declaration`: The schema definitions for tables and cells
 * - `schema`: A type-only placeholder (undefined at runtime) for TypeScript inference
 *
 * The actual working schema is only available through `engine.schema` after
 * attaching the definition to a FormulaEngine instance.
 *
 * @example
 * ```typescript
 * const mySchema = defineSchema<CellMetadata>()
 *   .addTableSchema("users", { workbookName: "wb1", tableName: "users" }, headers)
 *   .addCellSchema("config", cellAddress, parse);
 *
 * const engine = new FormulaEngine(mySchema);
 * engine.schema.users.findWhere({ id: 1 }); // Works!
 * engine.schema.config.read(); // Works!
 * mySchema.schema.users.findWhere({ id: 1 });  // Error: schema is undefined at runtime
 * ```
 */
export function defineSchema<
  TCellMetadata = unknown,
  TCurrentSchema extends Record<string, object> = Record<string, object>,
  TCurrentDeclaration extends Record<string, TableSchemaDefinition | CellSchemaDefinition> = Record<
    string,
    TableSchemaDefinition | CellSchemaDefinition
  >
>(): CreateSchema<TCellMetadata, TCurrentSchema, TCurrentDeclaration> {
  const declaration: Record<string, TableSchemaDefinition | CellSchemaDefinition> = {};

  const builder: CreateSchema<TCellMetadata, TCurrentSchema, TCurrentDeclaration> = {
    addTableSchema(namespace, address, headers) {
      declaration[namespace] = {
        type: "table",
        tableName: address.tableName,
        workbookName: address.workbookName,
        headers: headers as any,
      };
      return builder as any;
    },

    addCellSchema(namespace, cellAddress, parse) {
      declaration[namespace] = {
        type: "cell",
        cellAddress,
        parse: parse as any,
      };
      return builder as any;
    },

    // schema is undefined at runtime - it only exists for TypeScript type inference
    // The actual working schema is built by the FormulaEngine constructor
    schema: undefined as any as TCurrentSchema,

    declaration: declaration as TCurrentDeclaration,
  };

  return builder;
}

type ParseFunction<TCellMetadata> = (
  value: unknown,
  metadata: TCellMetadata
) => unknown;

export interface TableSchemaDefinition {
  type: "table";
  headers: Headers<unknown>;
  tableName: string;
  workbookName: string;
}

export interface CellSchemaDefinition {
  type: "cell";
  cellAddress: CellAddress;
  parse: (value: unknown, metadata: unknown) => unknown;
}

type Headers<TCellMetadata> = Record<
  string,
  {
    parse: ParseFunction<TCellMetadata>;
    index: number;
  }
>;

/**
 * Type representing the TableOrm methods exposed on the schema
 */
export type TableOrmSchema<TItem extends Record<string, unknown>> = {
  findWhere(filter: Partial<TItem>): TItem | undefined;
  findAllWhere(filter: Partial<TItem>): TItem[];
  append(item: TItem): TItem;
  updateWhere(filter: Partial<TItem>, update: Partial<TItem>): number;
  removeWhere(filter: Partial<TItem>): number;
  count(): number;
};

/**
 * Type representing the CellOrm methods exposed on the schema
 */
export type CellOrmSchema<TValue> = {
  read(): TValue;
  write(value: TValue): void;
  getAddress(): CellAddress;
};

export type SchemaDeclaration = Record<string, TableSchemaDefinition | CellSchemaDefinition>;
export type Schema = Record<string, object>;

export type CreateSchema<
  TCellMetadata,
  TCurrentSchema extends Schema,
  TCurrentDeclaration extends SchemaDeclaration
> = {
  addTableSchema<T extends string, THeaders extends Headers<TCellMetadata>>(
    namespace: T,
    address: { workbookName: string; tableName: string },
    headers: THeaders
  ): CreateSchema<
    TCellMetadata,
    TCurrentSchema & {
      [K in T]: TableOrmSchema<{
        [H in keyof THeaders]: ReturnType<THeaders[H]["parse"]>;
      }>;
    },
    TCurrentDeclaration & {
      [K in T]: {
        type: "table";
        tableName: string;
        workbookName: string;
        headers: THeaders;
      };
    }
  >;
  addCellSchema<T extends string, TValue>(
    namespace: T,
    cellAddress: CellAddress,
    parse: (value: unknown, metadata: TCellMetadata) => TValue
  ): CreateSchema<
    TCellMetadata,
    TCurrentSchema & {
      [K in T]: CellOrmSchema<TValue>;
    },
    TCurrentDeclaration & {
      [K in T]: {
        type: "cell";
        cellAddress: CellAddress;
        parse: (value: unknown, metadata: unknown) => unknown;
      };
    }
  >;
  schema: TCurrentSchema;
  declaration: TCurrentDeclaration;
};
