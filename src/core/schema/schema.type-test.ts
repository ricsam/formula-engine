import { FormulaEngine } from "../engine";
import { defineSchema } from "./schema";

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

const mySchema = defineSchema<CellMetadata>()
  .addTableSchema(
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
  .addCellSchema(
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

const schema = mySchema.schema;
const declaration = mySchema.declaration;
declaration.users.headers.email;

schema.project.read().linkTo;

schema.project.read();
schema.users.findWhere({ id: 1 });

type Assert<T, U extends T> = U;

type test1 = Assert<
  {
    value: string;
    linkTo: string | undefined;
  },
  ReturnType<typeof schema.project.read>
>;

type test2 = Assert<
  {
    id: number;
    name: string;
    email: string;
    age: number;
  },
  NonNullable<ReturnType<typeof schema.users.findWhere>>
>;

const engine = new FormulaEngine<
  {
    cell: CellMetadata;
  },
  typeof mySchema
>(mySchema);

type test3 = Assert<
  string,
  NonNullable<ReturnType<typeof engine.schema.project.read>["linkTo"]>
>;

engine.schema.project.read().linkTo;

const engineWithUndefinedSchema = new FormulaEngine();
type test4 = Assert<Record<string, object> | undefined, typeof engineWithUndefinedSchema.schema>;
