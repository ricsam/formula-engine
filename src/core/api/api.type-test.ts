import { FormulaEngine } from "../engine";
import { defineApi } from "./api";

type CellMetadata = { linkTo?: string };

const parseNumber = (value: unknown) => {
  if (typeof value !== "number") {
    throw new Error("Expected a number value");
  }
  return value;
};

const parseString = (value: unknown) => {
  if (typeof value !== "string") {
    throw new Error("Expected a string value");
  }
  return value;
};

const myApi = defineApi<CellMetadata>()
  .addTableApi(
    "users",
    {
      workbookName: "wb1",
      tableName: "users",
    },
    {
      id: {
        parse: (value, metadata) => {
          return parseNumber(value);
        },
        index: 0,
      },
      name: {
        parse: (value, metadata) => {
          return parseString(value);
        },
        index: 1,
      },
      email: {
        parse: (value, metadata) => {
          return parseString(value);
        },
        index: 2,
      },
      age: {
        parse: (value, metadata) => {
          return parseNumber(value);
        },
        index: 3,
      },
    }
  )
  .addCellApi(
    "project",
    {
      workbookName: "wb1",
      sheetName: "sheet1",
      colIndex: 0,
      rowIndex: 0,
    },
    (value, metadata) => {
      return {
        value: parseString(value),
        linkTo: metadata.linkTo,
      };
    }
  );

const api = myApi.api;
const declaration = myApi.declaration;
declaration.users.headers.email;

api.project.read().linkTo;

api.project.read();
api.users.findWhere({ id: 1 });

type Assert<T, U extends T> = U;

type test1 = Assert<
  {
    value: string;
    linkTo: string | undefined;
  },
  ReturnType<typeof api.project.read>
>;

type test2 = Assert<
  {
    id: number;
    name: string;
    email: string;
    age: number;
  },
  NonNullable<ReturnType<typeof api.users.findWhere>>
>;

const engine = new FormulaEngine<
  {
    cell: CellMetadata;
  },
  typeof myApi
>(myApi);

type test3 = Assert<
  string,
  NonNullable<ReturnType<typeof engine.api.project.read>["linkTo"]>
>;

engine.api.project.read().linkTo;

const engineWithUndefinedApi = new FormulaEngine();
type test4 = Assert<undefined, typeof engineWithUndefinedApi.api>;
