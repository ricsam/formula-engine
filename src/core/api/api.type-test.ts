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
    },
    {
      get(id: number) {
        return this.findWhere({ id });
      },
      create(newUser: { name: string; email: string; age: number }) {
        return this.append({
          id: this.count() + 1,
          ...newUser,
        });
      },
      update(update: {
        id: number;
        name?: string;
        email?: string;
        age?: number;
      }) {
        this.updateWhere({ id: update.id }, update);
      },
      delete(id: number) {
        this.removeWhere({ id });
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
    },
    {
      getCellValue() {
        return this.read();
      },
    }
  );

const api = myApi.api;
const declaration = myApi.declaration;
declaration.users.headers.email;

api.project.getCellValue().linkTo;

api.project.getCellValue();
api.users.get(1);

type Assert<T, U extends T> = U;

type test1 = Assert<
  {
    value: string;
    linkTo: string | undefined;
  },
  ReturnType<typeof api.project.getCellValue>
>;

type test2 = Assert<
  {
    id: number;
    name: string;
    email: string;
    age: number;
  },
  NonNullable<ReturnType<typeof api.users.get>>
>;

const engine = new FormulaEngine<
  {
    cell: CellMetadata;
  },
  typeof myApi
>(myApi);

type test3 = Assert<
  string,
  NonNullable<ReturnType<typeof engine.api.project.getCellValue>["linkTo"]>
>;

engine.api.project.getCellValue().linkTo;

const engineWithUndefinedApi = new FormulaEngine();
type test4 = Assert<undefined, typeof engineWithUndefinedApi.api>;
