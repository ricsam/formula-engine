import { FormulaEngine } from "../engine";
import type { SerializedCellValue } from "../types";
import { defineHeader, defineSchema } from "./schema";

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
      id: defineHeader(0, parseNumber),
      name: defineHeader(
        1,
        (value, metadata) => {
          if (typeof metadata.linkTo === "string") {
            return metadata.linkTo;
          }
          return parseString(metadata);
        },
        (value) =>
          value.startsWith("http")
            ? { value, metadata: { linkTo: value } }
            : { value }
      ),
      email: defineHeader(2, parseString),
      age: defineHeader(3, parseNumber),
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
    },
    (value) => ({
      value: value.value,
      metadata: value.linkTo ? { linkTo: value.linkTo } : undefined,
    })
  )
  .addGridSchema(
    "plate",
    { workbookName: "wb1", sheetName: "sheet1" },
    { start: { col: 0, row: 0 }, end: { col: 4, row: 10 } },
    (value, metadata) => {
      return parseNumber(value);
    },
    (value) => ({ value: value })
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

// Grid schema type tests
type test5 = Assert<readonly number[][], typeof schema.plate.columns>;
type test6 = Assert<readonly number[][], typeof schema.plate.rows>;
type test7 = Assert<number, (typeof schema.plate.rows)[1][4]>;

schema.plate.setValue(1, { col: 0, row: 0 });
schema.plate.getValue({ col: 0, row: 0 });

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

// Grid schema via engine
type test8 = Assert<readonly number[][], typeof engine.schema.plate.columns>;
type test9 = Assert<readonly number[][], typeof engine.schema.plate.rows>;

const engineWithUndefinedSchema = new FormulaEngine();
type test4 = Assert<
  Record<string, object> | undefined,
  typeof engineWithUndefinedSchema.schema
>;

// ============================================================================
// USAGE EXAMPLES
// ============================================================================

// ----------------------------------------------------------------------------
// Example 1: Simple case - parse returns SerializedCellValue (write optional)
// ----------------------------------------------------------------------------

// When parse returns number, string, or boolean, write is OPTIONAL
// because these types are directly serializable
const simpleSchema = defineSchema()
  // Cell returning number - write optional
  .addCellSchema(
    "counter",
    { workbookName: "wb", sheetName: "s", colIndex: 0, rowIndex: 0 },
    (value) => parseNumber(value)
    // No write function needed! ✓
  )
  // Cell returning string - write optional
  .addCellSchema(
    "title",
    { workbookName: "wb", sheetName: "s", colIndex: 1, rowIndex: 0 },
    (value) => parseString(value)
    // No write function needed! ✓
  )
  // Grid returning number - write optional
  .addGridSchema(
    "matrix",
    { workbookName: "wb", sheetName: "s" },
    { start: { col: 0, row: 0 }, end: { col: 9, row: 9 } },
    (value) => parseNumber(value)
    // No write function needed! ✓
  );

// Usage works seamlessly
const _counter: number = simpleSchema.schema.counter.read();
const _title: string = simpleSchema.schema.title.read();
const _matrix: readonly (readonly number[])[] = simpleSchema.schema.matrix.rows;

// ----------------------------------------------------------------------------
// Example 2: Complex case - parse returns object (write REQUIRED)
// ----------------------------------------------------------------------------

// When parse returns an object or complex type, write is REQUIRED
// because the engine needs to know how to serialize it back

type RichText = {
  text: string;
  bold?: boolean;
  italic?: boolean;
};

const complexSchema = defineSchema()
  // Cell returning an object - write IS REQUIRED
  .addCellSchema(
    "richContent",
    { workbookName: "wb", sheetName: "s", colIndex: 0, rowIndex: 0 },
    (value): RichText => ({
      text: parseString(value),
      bold: false,
    }),
    // Write function converts RichText back to SerializedCellValue
    (richText) => ({ value: richText.text })
  )
  // Grid returning an object - write IS REQUIRED
  .addGridSchema(
    "dataPoints",
    { workbookName: "wb", sheetName: "s" },
    { start: { col: 0, row: 0 }, end: { col: 4, row: 4 } },
    (value): { raw: number; formatted: string } => {
      const num = parseNumber(value);
      return { raw: num, formatted: `$${num.toFixed(2)}` };
    },
    // Write function extracts the raw value
    (point) => ({ value: point.raw })
  );

// Usage with type safety
const _rich: RichText = complexSchema.schema.richContent.read();
complexSchema.schema.richContent.write({ text: "Hello", bold: true });

const _points: readonly (readonly { raw: number; formatted: string }[])[] =
  complexSchema.schema.dataPoints.rows;
complexSchema.schema.dataPoints.setValue(
  { raw: 100, formatted: "$100.00" },
  { col: 0, row: 0 }
);

// ----------------------------------------------------------------------------
// Example 3: TypeScript ERROR when write is missing for complex types
// ----------------------------------------------------------------------------

// Uncomment the lines below to see TypeScript errors:
// When parse returns an object type, write is REQUIRED

/*
// ERROR: Missing write function for cell returning object
const _errorSchema1 = defineSchema().addCellSchema(
  "broken",
  { workbookName: "wb", sheetName: "s", colIndex: 0, rowIndex: 0 },
  (value): { custom: string } => ({ custom: parseString(value) })
  // ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  // Error: Argument of type '(value: unknown) => { custom: string; }'
  // is not assignable to parameter of type
  // '(value: unknown, metadata: unknown) => SerializedCellValue'.
  // Type '{ custom: string; }' is not assignable to type 'SerializedCellValue'.
);

// ERROR: Missing write function for grid returning object
const _errorSchema2 = defineSchema().addGridSchema(
  "brokenGrid",
  { workbookName: "wb", sheetName: "s" },
  { start: { col: 0, row: 0 }, end: { col: 2, row: 2 } },
  (value): { processed: number } => ({ processed: parseNumber(value) * 2 }),
  // ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  // Error: Argument of type '(value: unknown) => { processed: number; }'
  // is not assignable to parameter of type
  // '(value: unknown, metadata: unknown) => SerializedCellValue'.
  // Type '{ processed: number; }' is not assignable to type 'SerializedCellValue'.
);
*/

// ----------------------------------------------------------------------------
// Example 4: Table with headers - using defineHeader helper
// ----------------------------------------------------------------------------

type ProductMetadata = { sku?: string };

const productSchema = defineSchema<ProductMetadata>()
  .addTableSchema(
    "products",
    { workbookName: "inventory", tableName: "Products" },
    {
      // Simple columns - write is optional (returns SerializedCellValue)
      id: defineHeader(0, parseNumber),
      name: defineHeader(1, parseString),
      price: defineHeader(2, parseNumber),

      // Column with custom write - stores metadata
      code: defineHeader(
        3,
        // Parse reads the cell value and metadata
        (value, metadata) => parseString(value),
        // Write function converts back to serializable form with optional metadata
        (code) => ({
          value: code,
          // Could store sku in metadata if needed
        })
      ),
    }
  );

// Type-safe table operations
const _product = productSchema.schema.products.findWhere({ id: 1 });
if (_product) {
  // TypeScript knows the shape
  const _id: number = _product.id;
  const _name: string = _product.name;
  const _price: number = _product.price;
  const _code: string = _product.code;
}

productSchema.schema.products.append({
  id: 1,
  name: "Widget",
  price: 9.99,
  code: "WGT-001",
});
