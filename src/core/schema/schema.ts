import type {
  CellAddress,
  FiniteSpreadsheetRange,
  SerializedCellValue,
} from "../types";

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
  TCurrentDeclaration extends Record<
    string,
    TableSchemaDefinition | CellSchemaDefinition | GridSchemaDefinition
  > = Record<
    string,
    TableSchemaDefinition | CellSchemaDefinition | GridSchemaDefinition
  >
>(): CreateSchema<TCellMetadata, TCurrentSchema, TCurrentDeclaration> {
  const declaration: Record<
    string,
    TableSchemaDefinition | CellSchemaDefinition | GridSchemaDefinition
  > = {};

  const builder: CreateSchema<
    TCellMetadata,
    TCurrentSchema,
    TCurrentDeclaration
  > = {
    addTableSchema(namespace, address, headers) {
      declaration[namespace] = {
        type: "table",
        tableName: address.tableName,
        workbookName: address.workbookName,
        headers: headers as any,
      };
      return builder as any;
    },

    addCellSchema(
      namespace: string,
      cellAddress: CellAddress,
      parse: any,
      write?: any
    ) {
      declaration[namespace] = {
        type: "cell",
        cellAddress,
        parse,
        write: write ?? ((value: unknown) => ({ value: value as SerializedCellValue })),
      };
      return builder as any;
    },

    addGridSchema(
      namespace: string,
      address: { workbookName: string; sheetName: string },
      range: FiniteSpreadsheetRange,
      parse: any,
      write?: any
    ) {
      declaration[namespace] = {
        type: "grid",
        workbookName: address.workbookName,
        sheetName: address.sheetName,
        range,
        parse,
        write: write ?? ((value: unknown) => ({ value: value as SerializedCellValue })),
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
  write: (value: unknown) => {
    value: SerializedCellValue;
    metadata?: unknown;
  };
}

export interface GridSchemaDefinition {
  type: "grid";
  workbookName: string;
  sheetName: string;
  range: FiniteSpreadsheetRange;
  parse: (value: unknown, metadata: unknown) => unknown;
  write: (value: unknown) => {
    value: SerializedCellValue;
    metadata?: unknown;
  };
}

type Headers<TCellMetadata> = Record<
  string,
  {
    parse: (value: SerializedCellValue, metadata: TCellMetadata) => unknown;
    write: (value: any) => {
      value: SerializedCellValue;
      metadata?: TCellMetadata;
    };
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

/**
 * Type representing the GridOrm methods exposed on the schema
 */
export type GridOrmSchema<TValue> = {
  columns: readonly TValue[][];
  rows: readonly TValue[][];
  setValue(value: TValue, position: { col: number; row: number }): void;
  getValue(position: { col: number; row: number }): TValue;
};

export type SchemaDeclaration = Record<
  string,
  TableSchemaDefinition | CellSchemaDefinition | GridSchemaDefinition
>;
export type Schema = Record<string, object>;

/**
 * Write function type for converting parsed values back to serializable form
 */
type WriteFunction<TValue, TCellMetadata> = (value: TValue) => {
  value: SerializedCellValue;
  metadata?: TCellMetadata;
};

/**
 * Helper type to create the return type for addCellSchema/addGridSchema
 */
type AddCellSchemaResult<
  TCellMetadata,
  TCurrentSchema extends Schema,
  TCurrentDeclaration extends SchemaDeclaration,
  T extends string,
  TValue
> = CreateSchema<
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

type AddGridSchemaResult<
  TCellMetadata,
  TCurrentSchema extends Schema,
  TCurrentDeclaration extends SchemaDeclaration,
  T extends string,
  TValue
> = CreateSchema<
  TCellMetadata,
  TCurrentSchema & {
    [K in T]: GridOrmSchema<TValue>;
  },
  TCurrentDeclaration & {
    [K in T]: {
      type: "grid";
      workbookName: string;
      sheetName: string;
      range: FiniteSpreadsheetRange;
      parse: (value: unknown, metadata: unknown) => unknown;
    };
  }
>;

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

  // Overload 1: TValue is any type - write is required
  addCellSchema<T extends string, TValue>(
    namespace: T,
    cellAddress: CellAddress,
    parse: (value: unknown, metadata: TCellMetadata) => TValue,
    write: WriteFunction<TValue, TCellMetadata>
  ): AddCellSchemaResult<TCellMetadata, TCurrentSchema, TCurrentDeclaration, T, TValue>;

  // Overload 2: TValue extends SerializedCellValue - write is optional
  addCellSchema<T extends string, TValue extends SerializedCellValue>(
    namespace: T,
    cellAddress: CellAddress,
    parse: (value: unknown, metadata: TCellMetadata) => TValue,
    write?: WriteFunction<TValue, TCellMetadata>
  ): AddCellSchemaResult<TCellMetadata, TCurrentSchema, TCurrentDeclaration, T, TValue>;

  // Overload 1: TValue is any type - write is required
  addGridSchema<T extends string, TValue>(
    namespace: T,
    address: { workbookName: string; sheetName: string },
    range: FiniteSpreadsheetRange,
    parse: (value: unknown, metadata: TCellMetadata) => TValue,
    write: WriteFunction<TValue, TCellMetadata>
  ): AddGridSchemaResult<TCellMetadata, TCurrentSchema, TCurrentDeclaration, T, TValue>;

  // Overload 2: TValue extends SerializedCellValue - write is optional
  addGridSchema<T extends string, TValue extends SerializedCellValue>(
    namespace: T,
    address: { workbookName: string; sheetName: string },
    range: FiniteSpreadsheetRange,
    parse: (value: unknown, metadata: TCellMetadata) => TValue,
    write?: WriteFunction<TValue, TCellMetadata>
  ): AddGridSchemaResult<TCellMetadata, TCurrentSchema, TCurrentDeclaration, T, TValue>;

  schema: TCurrentSchema;
  declaration: TCurrentDeclaration;
};

export const createHeader = <TValue, TMetadata>(
  index: number,
  parse: (value: SerializedCellValue, metadata: TMetadata) => TValue,
  write?: (value: TValue) => {
    value: SerializedCellValue;
    metadata?: TMetadata;
  }
) => {
  return {
    parse,
    write: write ?? ((value) => ({ value })),
    index,
  };
};
