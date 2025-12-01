import type { CellAddress } from "../types";
import type { CellOrm } from "./cell-orm";
import type { TableOrm } from "./table-orm";

/**
 * Define an API schema for the FormulaEngine.
 *
 * The returned object contains:
 * - `declaration`: The schema definitions for tables and cells
 * - `api`: A type-only placeholder (undefined at runtime) for TypeScript inference
 *
 * The actual working API is only available through `engine.api` after
 * attaching the definition to a FormulaEngine instance.
 *
 * @example
 * ```typescript
 * const myApi = defineApi<CellMetadata>()
 *   .addTableApi("users", { workbookName: "wb1", tableName: "users" }, headers, methods)
 *   .addCellApi("config", cellAddress, parse, methods);
 *
 * const engine = new FormulaEngine(myApi);
 * engine.api.users.get(1); // Works!
 * myApi.api.users.get(1);  // Error: api is undefined at runtime
 * ```
 */
export function defineApi<
  TCellMetadata = unknown,
  TCurrentApi extends Record<
    string,
    Record<string, (...args: any[]) => any>
  > = Record<string, Record<string, (...args: any[]) => any>>,
  TCurrentDeclaration extends Record<string, TableApi | CellApi> = Record<
    string,
    TableApi | CellApi
  >
>(): CreateApi<TCellMetadata, TCurrentApi, TCurrentDeclaration> {
  const declaration: Record<string, TableApi | CellApi> = {};

  const builder: CreateApi<TCellMetadata, TCurrentApi, TCurrentDeclaration> = {
    addTableApi(namespace, address, headers, methods) {
      declaration[namespace] = {
        type: "table",
        tableName: address.tableName,
        workbookName: address.workbookName,
        headers: headers as any,
        methods: methods as any,
      };
      return builder as any;
    },

    addCellApi(namespace, cellAddress, parse, methods) {
      declaration[namespace] = {
        type: "cell",
        cellAddress,
        parse: parse as any,
        methods: methods as any,
      };
      return builder as any;
    },

    // api is undefined at runtime - it only exists for TypeScript type inference
    // The actual working API is built by the FormulaEngine constructor
    api: undefined as any as TCurrentApi,

    declaration: declaration as TCurrentDeclaration,
  };

  return builder;
}

// Keep createApi as an alias for backwards compatibility
export const createApi = defineApi;

type ParseFunction<TCellMetadata> = (
  value: unknown,
  metadata: TCellMetadata
) => unknown;

export interface TableApi {
  type: "table";
  headers: Headers<unknown>;
  tableName: string;
  workbookName: string;
  methods: TableMethods<Record<string, unknown>>;
}

export interface CellApi {
  type: "cell";
  cellAddress: CellAddress;
  parse: (value: unknown, metadata: unknown) => unknown;
  methods: CellMethods<unknown>;
}

type Headers<TCellMetadata> = Record<
  string,
  {
    parse: ParseFunction<TCellMetadata>;
    index: number;
  }
>;

type TableMethods<TItem extends Record<string, unknown>> = Record<
  string,
  (this: TableOrm<TItem>, ...args: any[]) => any
>;

type CellMethods<TValue> = Record<
  string,
  (this: CellOrm<TValue>, ...args: any[]) => any
>;

export type Declaration = Record<string, TableApi | CellApi>;
export type Api = Record<string, Record<string, (...args: any[]) => any>>;

export type CreateApi<
  TCellMetadata,
  TCurrentApi extends Api,
  TCurrentDeclaration extends Declaration
> = {
  addTableApi<
    T extends string,
    THeaders extends Headers<TCellMetadata>,
    TMethods extends TableMethods<{
      [K in keyof THeaders]: ReturnType<THeaders[K]["parse"]>;
    }>
  >(
    namespace: T,
    address: { workbookName: string; tableName: string },
    headers: THeaders,
    methods: TMethods
  ): CreateApi<
    TCellMetadata,
    TCurrentApi & {
      [K in T]: {
        [M in keyof TMethods]: (
          ...args: Parameters<TMethods[M]>
        ) => ReturnType<TMethods[M]>;
      };
    },
    TCurrentDeclaration & {
      [K in T]: {
        type: "table";
        tableName: string;
        workbookName: string;
        headers: THeaders;
        methods: TMethods;
      };
    }
  >;
  addCellApi<T extends string, TValue, TMethods extends CellMethods<TValue>>(
    namespace: T,
    cellAddress: CellAddress,
    parse: (value: unknown, metadata: TCellMetadata) => TValue,
    methods: TMethods
  ): CreateApi<
    TCellMetadata,
    TCurrentApi & {
      [K in T]: {
        [M in keyof TMethods]: (
          ...args: Parameters<TMethods[M]>
        ) => ReturnType<TMethods[M]>;
      };
    },
    TCurrentDeclaration & {
      [K in T]: {
        type: "cell";
        cellAddress: CellAddress;
        parse: (value: unknown, metadata: unknown) => unknown;
        methods: TMethods;
      };
    }
  >;
  api: TCurrentApi;
  declaration: TCurrentDeclaration;
};
