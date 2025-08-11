import { test, expect, describe } from "bun:test";
import { parseFormula } from "../../src/parser/parser";
import { Evaluator } from "../../src/evaluator/evaluator";
import { DependencyGraph } from "../../src/evaluator/dependency-graph";
import { ErrorHandler } from "../../src/evaluator/error-handler";
import type {
  EvaluationContext,
  FunctionDefinition,
  FunctionEvaluationResult,
} from "../../src/evaluator/evaluator";
import type {
  CellValue,
  SimpleCellAddress,
  SimpleCellRange,
} from "../../src/core/types";

describe("Evaluator Integration Tests", () => {
  // Create a mock spreadsheet
  const spreadsheet = new Map<string, CellValue>([
    ["0:0:0", 10], // A1 = 10
    ["0:1:0", 20], // B1 = 20
    ["0:2:0", 30], // C1 = 30
    ["0:0:1", 100], // A2 = 100
    ["0:1:1", 200], // B2 = 200
    ["0:2:1", 300], // C2 = 300
    ["0:0:2", "Hello"], // A3 = "Hello"
    ["0:1:2", "World"], // B3 = "World"
    ["0:0:3", true], // A4 = TRUE
    ["0:1:3", false], // B4 = FALSE
  ]);

  // Create function library
  const functions = new Map<string, FunctionDefinition>();

  // Basic math functions
  functions.set("SUM", {
    name: "SUM",
    minArgs: 1,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      let sum = 0;
      for (const arg of args) {
        if (Array.isArray(arg)) {
          const flat = arg.flat();
          for (const val of flat) {
            if (typeof val === "number") sum += val;
          }
        } else if (typeof arg === "number") {
          sum += arg;
        }
      }
      return {
        type: "value",
        value: sum,
      };
    },
  });

  functions.set("AVERAGE", {
    name: "AVERAGE",
    minArgs: 1,
    evaluate: ({ flatArgValues: args }):  FunctionEvaluationResult => {
      let sum = 0;
      let count = 0;
      for (const arg of args) {
        if (Array.isArray(arg)) {
          const flat = arg.flat();
          for (const val of flat) {
            if (typeof val === "number") {
              sum += val;
              count++;
            }
          }
        } else if (typeof arg === "number") {
          sum += arg;
          count++;
        }
      }
      return {
        type: "value",
        value: count > 0 ? sum / count : "#DIV/0!",
      };
    },
  });

  functions.set("MAX", {
    name: "MAX",
    minArgs: 1,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      let max = -Infinity;
      let hasNumbers = false;
      for (const arg of args) {
        if (Array.isArray(arg)) {
          const flat = arg.flat();
          for (const val of flat) {
            if (typeof val === "number") {
              max = Math.max(max, val);
              hasNumbers = true;
            }
          }
        } else if (typeof arg === "number") {
          max = Math.max(max, arg);
          hasNumbers = true;
        }
      }
      return {
        type: "value",
        value: hasNumbers ? max : 0,
      };
    },
  });

  functions.set("MIN", {
    name: "MIN",
    minArgs: 1,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      let min = Infinity;
      let hasNumbers = false;
      for (const arg of args) {
        if (Array.isArray(arg)) {
          const flat = arg.flat();
          for (const val of flat) {
            if (typeof val === "number") {
              min = Math.min(min, val);
              hasNumbers = true;
            }
          }
        } else if (typeof arg === "number") {
          min = Math.min(min, arg);
          hasNumbers = true;
        }
      }
      return {
        type: "value",
        value: hasNumbers ? min : 0,
      };
    },
  });

  // Text functions
  functions.set("CONCATENATE", {
    name: "CONCATENATE",
    minArgs: 1,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      let result = "";
      for (const arg of args) {
        if (Array.isArray(arg)) {
          const flat = arg.flat();
          for (const val of flat) {
            result += String(val ?? "");
          }
        } else {
          result += String(arg ?? "");
        }
      }
      return {
        type: "value",
        value: result,
      };
    },
  });

  functions.set("UPPER", {
    name: "UPPER",
    minArgs: 1,
    maxArgs: 1,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      const text = args[0];
      if (typeof text !== "string") return { type: "value", value: "#VALUE!" };
      return { type: "value", value: text.toUpperCase() };
    },
  });

  // Logical functions
  functions.set("IF", {
    name: "IF",
    minArgs: 3,
    maxArgs: 3,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      const condition = args[0];
      const trueValue = args[1];
      const falseValue = args[2];

      // Coerce to boolean
      let bool = false;
      if (typeof condition === "boolean") bool = condition;
      else if (typeof condition === "number") bool = condition !== 0;
      else if (typeof condition === "string") bool = condition.length > 0;

      return {
        type: "value",
        value: bool ? trueValue : falseValue,
      };
    },
  });

  functions.set("AND", {
    name: "AND",
    minArgs: 1,
    evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
      for (const arg of args) {
        if (Array.isArray(arg)) {
          const flat = arg.flat();
          for (const val of flat) {
            if (!val) return { type: "value", value: false };
          }
        } else {
          if (!arg) return { type: "value", value: false };
        }
      }
      return {
        type: "value",
        value: true,
      };
    },
  });

  // Create evaluation context
  function createContext(): EvaluationContext {
    const dependencyGraph = new DependencyGraph();
    const errorHandler = new ErrorHandler();

    return {
      currentSheet: 0,
      currentCell: undefined,
      namedExpressions: new Map([
        ["PI", { name: "PI", expression: "3.14159" }],
        ["TaxRate", { name: "TaxRate", expression: "0.08" }],
      ]),
      getCellValue: (address: SimpleCellAddress) => {
        const key = `${address.sheet}:${address.col}:${address.row}`;
        return spreadsheet.get(key) ?? undefined;
      },
      getRangeValues: (range: SimpleCellRange) => {
        const result: CellValue[][] = [];
        for (let row = range.start.row; row <= range.end.row; row++) {
          const rowData: CellValue[] = [];
          for (let col = range.start.col; col <= range.end.col; col++) {
            const key = `${range.start.sheet}:${col}:${row}`;
            rowData.push(spreadsheet.get(key) ?? undefined);
          }
          result.push(rowData);
        }
        return result;
      },
      getFunction: (name: string) => functions.get(name),
      errorHandler,
      evaluationStack: new Set(),
    };
  }

  function evaluateFormula(formula: string): CellValue | CellValue[][] {
    try {
      const ast = parseFormula(formula);

      const context = createContext();
      const dependencyGraph = new DependencyGraph();
      const errorHandler = new ErrorHandler();
      const evaluator = new Evaluator(dependencyGraph, functions, errorHandler);

      const result = evaluator.evaluate(ast, context);
      return result.value;
    } catch (error) {
      console.log(error);
      return "#ERROR!";
    }
  }

  describe("Basic calculations", () => {
    test("should evaluate simple arithmetic", () => {
      expect(evaluateFormula("=1+2")).toBe(3);
      expect(evaluateFormula("=10-5")).toBe(5);
      expect(evaluateFormula("=4*5")).toBe(20);
      expect(evaluateFormula("=20/4")).toBe(5);
      expect(evaluateFormula("=2^3")).toBe(8);
    });

    test("should follow operator precedence", () => {
      expect(evaluateFormula("=2+3*4")).toBe(14);
      expect(evaluateFormula("=(2+3)*4")).toBe(20);
      expect(evaluateFormula("=2^3*4")).toBe(32);
      expect(evaluateFormula("=100/5/2")).toBe(10);
    });

    test("should handle negation and percent", () => {
      expect(evaluateFormula("=-10")).toBe(-10);
      expect(evaluateFormula("=+10")).toBe(10);
      expect(evaluateFormula("=50%")).toBe(0.5);
      expect(evaluateFormula("=100*50%")).toBe(50);
    });
  });

  describe("Cell references", () => {
    test("should evaluate single cell references", () => {
      expect(evaluateFormula("=A1")).toBe(10);
      expect(evaluateFormula("=B1")).toBe(20);
      expect(evaluateFormula("=C1")).toBe(30);
      expect(evaluateFormula("=A3")).toBe("Hello");
    });

    test("should evaluate calculations with cell references", () => {
      expect(evaluateFormula("=A1+B1")).toBe(30);
      expect(evaluateFormula("=A1*B1")).toBe(200);
      expect(evaluateFormula("=A2/A1")).toBe(10);
      expect(evaluateFormula('=A3&" "&B3')).toBe("Hello World");
    });

    test("should handle empty cells", () => {
      expect(evaluateFormula("=D1")).toBeUndefined();
      expect(evaluateFormula("=D1+10")).toBe(10); // Empty cell treated as 0
      expect(evaluateFormula('=D1&"text"')).toBe("text"); // Empty cell treated as ""
    });
  });

  describe("Functions with ranges", () => {
    test("should evaluate SUM function", () => {
      expect(evaluateFormula("=SUM(A1:C1)")).toBe(60);
      expect(evaluateFormula("=SUM(A1:C2)")).toBe(660);
      expect(evaluateFormula("=SUM(A1,B1,C1)")).toBe(60);
      expect(evaluateFormula("=SUM(A1:C1,A2:C2)")).toBe(660);
    });

    test("should evaluate AVERAGE function", () => {
      expect(evaluateFormula("=AVERAGE(A1:C1)")).toBe(20);
      expect(evaluateFormula("=AVERAGE(A1:C2)")).toBe(110);
      expect(evaluateFormula("=AVERAGE(10,20,30)")).toBe(20);
    });

    test("should evaluate MAX function", () => {
      expect(evaluateFormula("=MAX(A1:C1)")).toBe(30);
      expect(evaluateFormula("=MAX(A1:C2)")).toBe(300);
      expect(evaluateFormula("=MAX(A1,B2,50)")).toBe(200);
    });
  });

  describe("Text functions", () => {
    test("should evaluate CONCATENATE", () => {
      expect(evaluateFormula('=CONCATENATE("Hello"," ","World")')).toBe(
        "Hello World"
      );
      expect(evaluateFormula("=CONCATENATE(A3,B3)")).toBe("HelloWorld");
      expect(evaluateFormula('=CONCATENATE(A3," ",B3)')).toBe("Hello World");
    });

    test("should evaluate UPPER", () => {
      expect(evaluateFormula('=UPPER("hello")')).toBe("HELLO");
      expect(evaluateFormula("=UPPER(A3)")).toBe("HELLO");
      expect(evaluateFormula("=UPPER(123)")).toBe("#VALUE!");
    });
  });

  describe("Logical functions", () => {
    test("should evaluate IF function", () => {
      expect(evaluateFormula('=IF(TRUE,"Yes","No")')).toBe("Yes");
      expect(evaluateFormula('=IF(FALSE,"Yes","No")')).toBe("No");
      expect(evaluateFormula('=IF(A1>5,"Greater","Less")')).toBe("Greater");
      expect(evaluateFormula('=IF(A1>50,"Greater","Less")')).toBe("Less");
    });

    test("should evaluate AND function", () => {
      expect(evaluateFormula("=AND(TRUE,TRUE)")).toBe(true);
      expect(evaluateFormula("=AND(TRUE,FALSE)")).toBe(false);
      expect(evaluateFormula("=AND(A1>5,B1>15)")).toBe(true);
      expect(evaluateFormula("=AND(A1>5,B1>25)")).toBe(false);
    });

    test("should evaluate complex conditions", () => {
      expect(evaluateFormula('=IF(AND(A1>5,B1<30),"Pass","Fail")')).toBe(
        "Pass"
      );
      expect(evaluateFormula("=IF(A1+B1>25,MAX(A1,B1),MIN(A1,B1))")).toBe(20);
    });
  });

  describe("Array operations", () => {
    test("should evaluate array literals", () => {
      // Array formulas return arrays which should be handled by the engine for spilling
      const result1 = evaluateFormula("={1,2,3}");
      expect(result1).toEqual([[1, 2, 3]]);

      const result2 = evaluateFormula("={1;2;3}");
      expect(result2).toEqual([[1], [2], [3]]);

      const result3 = evaluateFormula("={1,2;3,4}");
      expect(result3).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    test("should evaluate array formulas", () => {
      const result1 = evaluateFormula("=A1:C1+10") as unknown as number[][];
      expect(result1).toEqual([[20, 30, 40]]);

      const result2 = evaluateFormula("=A1:C1*2") as unknown as number[][];
      expect(result2).toEqual([[20, 40, 60]]);

      const result3 = evaluateFormula(
        "={1,2,3}+{10,20,30}"
      ) as unknown as number[][];
      expect(result3).toEqual([[11, 22, 33]]);
    });

    test("should handle array broadcasting", () => {
      const result1 = evaluateFormula("=A1:C1+100") as unknown as number[][];
      expect(result1).toEqual([[110, 120, 130]]);

      const result2 = evaluateFormula("={1;2;3}*10") as unknown as number[][];
      expect(result2).toEqual([[10], [20], [30]]);
    });
  });

  describe("Named expressions", () => {
    test("should evaluate named expressions", () => {
      expect(evaluateFormula("=PI")).toBe(3.14159);
      expect(evaluateFormula("=TaxRate")).toBe(0.08);
      expect(evaluateFormula("=100*TaxRate")).toBe(8);
      expect(evaluateFormula("=2*PI")).toBe(6.28318);
    });

    test("should handle unknown names", () => {
      expect(evaluateFormula("=UnknownName")).toBe("#NAME?");
    });
  });

  describe("Error handling", () => {
    test("should propagate errors", () => {
      expect(evaluateFormula("=1/0")).toBe("#DIV/0!");
      expect(evaluateFormula("=10+1/0")).toBe("#DIV/0!");
      expect(evaluateFormula("=SUM(A1,1/0,B1)")).toBe("#DIV/0!");
    });

    test("should handle type errors", () => {
      expect(evaluateFormula('="text"+10')).toBe("#VALUE!");
      expect(evaluateFormula("=UPPER(123)")).toBe("#VALUE!");
    });

    test("should handle function errors", () => {
      expect(evaluateFormula("=UNKNOWN()")).toBe("#NAME?");
      expect(evaluateFormula("=SUM()")).toBe("#ERROR!"); // Too few args - caught by parser
      expect(evaluateFormula('=UPPER("a","b")')).toBe("#ERROR!"); // Too many args - caught by parser
    });
  });

  describe("Complex formulas", () => {
    test("should evaluate nested functions", () => {
      expect(evaluateFormula("=SUM(MAX(A1,B1),MIN(A1,B1))")).toBe(30);
      expect(evaluateFormula("=IF(SUM(A1:C1)>50,AVERAGE(A1:C1),0)")).toBe(20);
      expect(evaluateFormula('=CONCATENATE("Total: ",SUM(A1:C1))')).toBe(
        "Total: 60"
      );
    });

    test("should evaluate complex calculations", () => {
      expect(evaluateFormula("=(SUM(A1:C1)+SUM(A2:C2))/6")).toBe(110);
      expect(evaluateFormula("=MAX(A1:C1)*MIN(A2:C2)")).toBe(3000);
      expect(evaluateFormula("=IF(AVERAGE(A1:C1)>15,SUM(A2:C2),0)")).toBe(600);
    });

    test("should handle mixed types correctly", () => {
      expect(evaluateFormula('=A1&" items"')).toBe("10 items");
      expect(evaluateFormula("=IF(A4,A1,B1)")).toBe(10); // A4 is TRUE
      expect(evaluateFormula("=IF(B4,A1,B1)")).toBe(20); // B4 is FALSE
    });
  });

  describe("Dependency tracking", () => {
    test("should track cell dependencies", () => {
      try {
        const ast = parseFormula("=A1+B1*C1");

        const context = createContext();
        const dependencyGraph = new DependencyGraph();
        const errorHandler = new ErrorHandler();
        const evaluator = new Evaluator(
          dependencyGraph,
          functions,
          errorHandler
        );

        const result = evaluator.evaluate(ast, context);

        expect(result.value).toBe(610); // 10 + 20*30
        expect(result.dependencies.has("0:0:0")).toBe(true); // A1
        expect(result.dependencies.has("0:1:0")).toBe(true); // B1
        expect(result.dependencies.has("0:2:0")).toBe(true); // C1
        expect(result.dependencies.size).toBe(3);
      } catch (error) {
        throw error;
      }
    });

    test("should track range dependencies", () => {
      try {
        const ast = parseFormula("=SUM(A1:C2)");

        const context = createContext();
        const dependencyGraph = new DependencyGraph();
        const errorHandler = new ErrorHandler();
        const evaluator = new Evaluator(
          dependencyGraph,
          functions,
          errorHandler
        );

        const result = evaluator.evaluate(ast, context);

        expect(result.value).toBe(660);
        expect(result.dependencies.has("0:0:0:2:1")).toBe(true); // Range A1:C2
        expect(result.dependencies.has("0:0:0")).toBe(true); // A1
        expect(result.dependencies.has("0:1:0")).toBe(true); // B1
        expect(result.dependencies.has("0:2:0")).toBe(true); // C1
        expect(result.dependencies.has("0:0:1")).toBe(true); // A2
        expect(result.dependencies.has("0:1:1")).toBe(true); // B2
        expect(result.dependencies.has("0:2:1")).toBe(true); // C2
      } catch (error) {
        throw error;
      }
    });
  });

  describe("Edge cases", () => {
    test("should handle empty formulas", () => {
      expect(evaluateFormula("=")).toBeUndefined();
    });

    test("should handle whitespace", () => {
      expect(evaluateFormula("= A1 + B1 ")).toBe(30);
      expect(evaluateFormula("=SUM( A1 : C1 )")).toBe(60);
    });

    test("should handle case insensitivity for functions", () => {
      expect(evaluateFormula("=sum(A1:C1)")).toBe(60);
      expect(evaluateFormula("=Sum(A1:C1)")).toBe(60);
      expect(evaluateFormula("=SUM(A1:C1)")).toBe(60);
    });
  });
});
