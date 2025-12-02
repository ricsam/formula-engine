import type { CellAddress } from "../types";

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
 *   .addTableApi("users", { workbookName: "wb1", tableName: "users" }, headers)
 *   .addCellApi("config", cellAddress, parse);
 *
 * const engine = new FormulaEngine(myApi);
 * engine.api.users.findWhere({ id: 1 }); // Works!
 * engine.api.config.read(); // Works!
 * myApi.api.users.findWhere({ id: 1 });  // Error: api is undefined at runtime
 * ```
 */
export function defineApi<
  TCellMetadata = unknown,
  TCurrentApi extends Record<string, object> = Record<string, object>,
  TCurrentDeclaration extends Record<string, TableApi | CellApi> = Record<
    string,
    TableApi | CellApi
  >
>(): CreateApi<TCellMetadata, TCurrentApi, TCurrentDeclaration> {
  const declaration: Record<string, TableApi | CellApi> = {};

  const builder: CreateApi<TCellMetadata, TCurrentApi, TCurrentDeclaration> = {
    addTableApi(namespace, address, headers) {
      declaration[namespace] = {
        type: "table",
        tableName: address.tableName,
        workbookName: address.workbookName,
        headers: headers as any,
      };
      return builder as any;
    },

    addCellApi(namespace, cellAddress, parse) {
      declaration[namespace] = {
        type: "cell",
        cellAddress,
        parse: parse as any,
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
}

export interface CellApi {
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
 * Type representing the TableOrm methods exposed on the API
 */
export type TableOrmApi<TItem extends Record<string, unknown>> = {
  findWhere(filter: Partial<TItem>): TItem | undefined;
  findAllWhere(filter: Partial<TItem>): TItem[];
  append(item: TItem): TItem;
  updateWhere(filter: Partial<TItem>, update: Partial<TItem>): number;
  removeWhere(filter: Partial<TItem>): number;
  count(): number;
};

/**
 * Type representing the CellOrm methods exposed on the API
 */
export type CellOrmApi<TValue> = {
  read(): TValue;
  write(value: TValue): void;
  getAddress(): CellAddress;
};

export type Declaration = Record<string, TableApi | CellApi>;
export type Api = Record<string, object>;

export type CreateApi<
  TCellMetadata,
  TCurrentApi extends Api,
  TCurrentDeclaration extends Declaration
> = {
  addTableApi<T extends string, THeaders extends Headers<TCellMetadata>>(
    namespace: T,
    address: { workbookName: string; tableName: string },
    headers: THeaders
  ): CreateApi<
    TCellMetadata,
    TCurrentApi & {
      [K in T]: TableOrmApi<{
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
  addCellApi<T extends string, TValue>(
    namespace: T,
    cellAddress: CellAddress,
    parse: (value: unknown, metadata: TCellMetadata) => TValue
  ): CreateApi<
    TCellMetadata,
    TCurrentApi & {
      [K in T]: CellOrmApi<TValue>;
    },
    TCurrentDeclaration & {
      [K in T]: {
        type: "cell";
        cellAddress: CellAddress;
        parse: (value: unknown, metadata: unknown) => unknown;
      };
    }
  >;
  api: TCurrentApi;
  declaration: TCurrentDeclaration;
};
