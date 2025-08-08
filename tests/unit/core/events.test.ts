import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SimpleCellAddress } from "../../../src/core/types";

describe("FormulaEngine Events System", () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("TestSheet");
    sheetId = engine.getSheetId(sheetName);
  });

  describe("onCellUpdate", () => {
    test("should notify when setting a value", () => {
      const events: any[] = [];
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const unsubscribe = engine.onCellUpdate(address, (event) => events.push(event));
      engine.setCellContent(address, 42);
      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({ address, oldValue: undefined, newValue: 42 });
      unsubscribe();
    });

    test("should notify when updating and clearing a value", () => {
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);
      const events: any[] = [];
      const unsubscribe = engine.onCellUpdate(address, (event) => events.push(event));
      engine.setCellContent(address, 84);
      engine.setCellContent(address, undefined);
      expect(events).toEqual([
        { address, oldValue: 42, newValue: 84 },
        { address, oldValue: 84, newValue: undefined },
      ]);
      unsubscribe();
    });

    test("should notify for formula cells", () => {
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const events: any[] = [];
      const unsubscribe = engine.onCellUpdate(address, (event) => events.push(event));
      engine.setCellContent(address, "=1+2");
      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({ address, oldValue: undefined, newValue: 3 });
      unsubscribe();
    });

    test("should not notify when value does not change", () => {
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);
      const events: any[] = [];
      const unsubscribe = engine.onCellUpdate(address, (event) => events.push(event));
      engine.setCellContent(address, 42);
      expect(events).toHaveLength(0);
      unsubscribe();
    });

    test("should support multiple listeners for the same cell and unsubscribe", () => {
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const events1: any[] = [];
      const events2: any[] = [];
      const unsub1 = engine.onCellUpdate(address, (e) => events1.push(e));
      const unsub2 = engine.onCellUpdate(address, (e) => events2.push(e));
      engine.setCellContent(address, 42);
      expect(events1).toHaveLength(1);
      expect(events2).toHaveLength(1);
      unsub1();
      engine.setCellContent(address, 84);
      expect(events1).toHaveLength(1);
      expect(events2).toHaveLength(2);
      unsub2();
    });
  });

  describe("onCellsUpdate", () => {
    test("should deliver batched updates for setSheetContent", () => {
      const batches: any[][] = [];
      const unsubscribe = engine.onCellsUpdate(sheetId, (events) => batches.push(events));
      const cellData = new Map([
        ["A1", 1],
        ["B1", 2],
        ["C1", 3],
      ]);
      engine.setSheetContent(sheetId, cellData);
      expect(batches.length).toBe(1);
      const batch = batches[0]!;
      const addressA1 = engine.simpleCellAddressFromString("A1", sheetId);
      const addressB1 = engine.simpleCellAddressFromString("B1", sheetId);
      const addressC1 = engine.simpleCellAddressFromString("C1", sheetId);
      expect(batch).toEqual(
        expect.arrayContaining([
          { address: addressA1, oldValue: undefined, newValue: 1 },
          { address: addressB1, oldValue: undefined, newValue: 2 },
          { address: addressC1, oldValue: undefined, newValue: 3 },
        ])
      );
      unsubscribe();
    });

    test("should notify only for the target sheet", () => {
      const otherSheet = engine.getSheetId(engine.addSheet("Other"));
      const events: any[][] = [];
      const unsubscribe = engine.onCellsUpdate(sheetId, (batch) => events.push(batch));
      engine.setCellContent({ sheet: otherSheet, col: 0, row: 0 }, 5);
      expect(events).toHaveLength(0);
      engine.setCellContent({ sheet: sheetId, col: 0, row: 0 }, 7);
      expect(events.length).toBeGreaterThan(0);
      unsubscribe();
    });

    test("should deliver full chain updates for A1->B1->C1->D1 when A1 changes", () => {
      // Setup chain: B1=A1+1, C1=B1+1, D1=C1+1
      engine.setCellContent({ sheet: sheetId, col: 1, row: 0 }, "=A1+1");
      engine.setCellContent({ sheet: sheetId, col: 2, row: 0 }, "=B1+1");
      engine.setCellContent({ sheet: sheetId, col: 3, row: 0 }, "=C1+1");

      const batches: any[][] = [];
      const unsubscribe = engine.onCellsUpdate(sheetId, (events) => batches.push(events));

      // Change A1 triggers the chain
      engine.setCellContent({ sheet: sheetId, col: 0, row: 0 }, 5);

      // Verify values are updated immediately after setCellContent
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 0 })).toBe(8);

      // Verify batch contains updates for A1, B1, C1, D1
      expect(batches.length).toBeGreaterThan(0);
      const allEvents = batches.flat();

      const hasA1 = allEvents.some((e) => e.address.col === 0 && e.address.row === 0 && e.newValue === 5);
      const hasB1 = allEvents.some((e) => e.address.col === 1 && e.address.row === 0 && e.newValue === 6);
      const hasC1 = allEvents.some((e) => e.address.col === 2 && e.address.row === 0 && e.newValue === 7);
      const hasD1 = allEvents.some((e) => e.address.col === 3 && e.address.row === 0 && e.newValue === 8);
      expect(hasA1 && hasB1 && hasC1 && hasD1).toBe(true);

      unsubscribe();
    });
  });

  describe("sheet events", () => {
    test("should emit sheet-added/removed/renamed events", () => {
      const added: any[] = [];
      const removed: any[] = [];
      const renamed: any[] = [];
      const un1 = engine.on("sheet-added", (e) => added.push(e));
      const un2 = engine.on("sheet-removed", (e) => removed.push(e));
      const un3 = engine.on("sheet-renamed", (e) => renamed.push(e));
      const name2 = engine.addSheet("NewSheet");
      const id2 = engine.getSheetId(name2);
      expect(added).toHaveLength(1);
      engine.renameSheet(id2, "Renamed");
      expect(renamed).toHaveLength(1);
      engine.removeSheet(id2);
      expect(removed).toHaveLength(1);
      un1();
      un2();
      un3();
    });
  });

  describe("event data & timing", () => {
    test("should provide correct addresses for multiple cells", () => {
      const a1: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const b6: SimpleCellAddress = { sheet: sheetId, col: 1, row: 5 };
      const z101: SimpleCellAddress = { sheet: sheetId, col: 25, row: 100 };
      const events: any[] = [];
      const u1 = engine.onCellUpdate(a1, (e) => events.push(e));
      const u2 = engine.onCellUpdate(b6, (e) => events.push(e));
      const u3 = engine.onCellUpdate(z101, (e) => events.push(e));
      engine.setCellContent(a1, 0);
      engine.setCellContent(b6, 1);
      engine.setCellContent(z101, 2);
      expect(events).toEqual(
        expect.arrayContaining([
          { address: a1, oldValue: undefined, newValue: 0 },
          { address: b6, oldValue: undefined, newValue: 1 },
          { address: z101, oldValue: undefined, newValue: 2 },
        ])
      );
      u1();
      u2();
      u3();
    });

    test("should handle formula evaluation in events", () => {
      const a1: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const b1: SimpleCellAddress = { sheet: sheetId, col: 1, row: 0 };
      const c1: SimpleCellAddress = { sheet: sheetId, col: 2, row: 0 };
      const events: any[] = [];
      const unsub = engine.onCellUpdate(c1, (e) => events.push(e));
      engine.setCellContent(a1, 10);
      engine.setCellContent(b1, 20);
      engine.setCellContent(c1, "=A1+B1");
      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({ address: c1, oldValue: undefined, newValue: 30 });
      unsub();
    });

    test("should fire cell updates synchronously", () => {
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      let fired = false;
      const unsub = engine.onCellUpdate(address, () => {
        fired = true;
      });
      engine.setCellContent(address, 42);
      expect(fired).toBe(true);
      unsub();
    });

    test("should include dependent recalculations in batch", () => {
      const events: { address: string; value: any }[] = [];
      const unsub = engine.onCellsUpdate(sheetId, (batch) => {
        for (const e of batch) {
          events.push({
            address: `${String.fromCharCode(65 + e.address.col)}${e.address.row + 1}`,
            value: e.newValue,
          });
        }
      });
      const a1: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const b1: SimpleCellAddress = { sheet: sheetId, col: 1, row: 0 };
      const c1: SimpleCellAddress = { sheet: sheetId, col: 2, row: 0 };
      engine.setCellContent(a1, 10);
      engine.setCellContent(b1, 20);
      engine.setCellContent(c1, "=A1+B1");
      events.length = 0;
      engine.setCellContent(a1, 15);
      expect(events).toEqual(
        expect.arrayContaining([
          { address: "A1", value: 15 },
          { address: "C1", value: 35 },
        ])
      );
      unsub();
    });
  });
});