import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../../src/core/engine';
import type { SimpleCellAddress } from '../../../src/core/types';

describe('FormulaEngine Events System', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    sheetId = engine.getSheetId(sheetName);
  });

  describe('cell-changed events', () => {
    test('should emit cell-changed event when setting a value', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        address,
        oldValue: undefined,
        newValue: 42
      });

      unsubscribe();
    });

    test('should emit cell-changed event when updating a value', () => {
      const events: any[] = [];
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      
      // Set initial value
      engine.setCellContent(address, 42);

      // Subscribe after initial set
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      // Update the value
      engine.setCellContent(address, 84);

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        address,
        oldValue: 42,
        newValue: 84
      });

      unsubscribe();
    });

    test('should emit cell-changed event when clearing a cell', () => {
      const events: any[] = [];
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      
      // Set initial value
      engine.setCellContent(address, 42);

      // Subscribe after initial set
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      // Clear the cell
      engine.setCellContent(address, undefined);

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        address,
        oldValue: 42,
        newValue: undefined
      });

      unsubscribe();
    });

    test('should emit cell-changed event for formula cells', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, '=1+2');

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        address,
        oldValue: undefined,
        newValue: 3
      });

      unsubscribe();
    });

    test('should emit multiple cell-changed events for multiple cells', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      const cellData = new Map([
        ['A1', 1],
        ['B1', 2],
        ['C1', 3]
      ]);

      engine.setSheetContent(sheetId, cellData);

      expect(events).toHaveLength(3);
      
      // Verify all events were emitted with correct values
      const addressA1 = engine.simpleCellAddressFromString('A1', sheetId);
      const addressB1 = engine.simpleCellAddressFromString('B1', sheetId);
      const addressC1 = engine.simpleCellAddressFromString('C1', sheetId);

      expect(events).toContainEqual({
        address: addressA1,
        oldValue: undefined,
        newValue: 1
      });
      expect(events).toContainEqual({
        address: addressB1,
        oldValue: undefined,
        newValue: 2
      });
      expect(events).toContainEqual({
        address: addressC1,
        oldValue: undefined,
        newValue: 3
      });

      unsubscribe();
    });

    test('should not emit cell-changed event when value does not change', () => {
      const events: any[] = [];
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      
      // Set initial value
      engine.setCellContent(address, 42);

      // Subscribe after initial set
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      // Set the same value again
      engine.setCellContent(address, 42);

      expect(events).toHaveLength(0);

      unsubscribe();
    });
  });

  describe('sheet-added events', () => {
    test('should emit sheet-added event when adding a sheet', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('sheet-added', (event) => {
        events.push(event);
      });

      const sheetName = engine.addSheet('NewSheet');
      const newSheetId = engine.getSheetId(sheetName);

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        sheetId: newSheetId,
        sheetName: 'NewSheet'
      });

      unsubscribe();
    });

    test('should emit sheet-added event with unique name when name conflicts', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('sheet-added', (event) => {
        events.push(event);
      });

      // Add sheet with existing name
      const sheetName = engine.addSheet('TestSheet');
      const newSheetId = engine.getSheetId(sheetName);

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        sheetId: newSheetId,
        sheetName: 'TestSheet_1' // Should be auto-renamed
      });

      unsubscribe();
    });
  });

  describe('sheet-removed events', () => {
    test('should emit sheet-removed event when removing a sheet', () => {
      // Add a second sheet first (can't remove the only sheet)
      const secondSheetName = engine.addSheet('SecondSheet');
      const secondSheetId = engine.getSheetId(secondSheetName);

      const events: any[] = [];
      const unsubscribe = engine.on('sheet-removed', (event) => {
        events.push(event);
      });

      engine.removeSheet(secondSheetId);

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        sheetId: secondSheetId,
        sheetName: 'SecondSheet'
      });

      unsubscribe();
    });

    test('should not emit sheet-removed event when removal is not possible', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('sheet-removed', (event) => {
        events.push(event);
      });

      // Try to remove the only sheet (should fail)
      engine.removeSheet(sheetId);

      expect(events).toHaveLength(0);

      unsubscribe();
    });
  });

  describe('sheet-renamed events', () => {
    test('should emit sheet-renamed event when renaming a sheet', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('sheet-renamed', (event) => {
        events.push(event);
      });

      engine.renameSheet(sheetId, 'RenamedSheet');

      expect(events).toHaveLength(1);
      expect(events[0]).toEqual({
        sheetId,
        oldName: 'TestSheet',
        newName: 'RenamedSheet'
      });

      unsubscribe();
    });

    test('should not emit sheet-renamed event when rename is not possible', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('sheet-renamed', (event) => {
        events.push(event);
      });

      // Try to rename to empty string (should fail)
      engine.renameSheet(sheetId, '');

      expect(events).toHaveLength(0);

      unsubscribe();
    });
  });

  describe('event subscription management', () => {
    test('should properly unsubscribe from events', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      // Set a value (should trigger event)
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);

      expect(events).toHaveLength(1);

      // Unsubscribe
      unsubscribe();

      // Set another value (should not trigger event)
      engine.setCellContent(address, 84);

      expect(events).toHaveLength(1); // Still only 1 event
    });

    test('should support multiple listeners for the same event', () => {
      const events1: any[] = [];
      const events2: any[] = [];
      
      const unsubscribe1 = engine.on('cell-changed', (event) => {
        events1.push(event);
      });
      
      const unsubscribe2 = engine.on('cell-changed', (event) => {
        events2.push(event);
      });

      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);

      expect(events1).toHaveLength(1);
      expect(events2).toHaveLength(1);
      expect(events1[0]).toEqual(events2[0]);

      unsubscribe1();
      unsubscribe2();
    });

    test('should support subscribe alias for on method', () => {
      const events: any[] = [];
      const unsubscribe = engine.subscribe('cell-changed', (event) => {
        events.push(event);
      });

      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);

      expect(events).toHaveLength(1);

      unsubscribe();
    });

    test('should remove all listeners with removeAllListeners', () => {
      const events1: any[] = [];
      const events2: any[] = [];
      
      engine.on('cell-changed', (event) => {
        events1.push(event);
      });
      
      engine.on('sheet-added', (event) => {
        events2.push(event);
      });

      // Remove all listeners
      engine.removeAllListeners();

      // Trigger events
      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);
      engine.addSheet('NewSheet');

      expect(events1).toHaveLength(0);
      expect(events2).toHaveLength(0);
    });
  });

  describe('event data integrity', () => {
    test('should provide correct address in cell-changed events', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      // Test different address positions
      const addresses = [
        { sheet: sheetId, col: 0, row: 0 },   // A1
        { sheet: sheetId, col: 1, row: 5 },   // B6
        { sheet: sheetId, col: 25, row: 100 } // Z101
      ];

      addresses.forEach((address, index) => {
        engine.setCellContent(address, index);
      });

      expect(events).toHaveLength(3);
      addresses.forEach((address, index) => {
        expect(events[index].address).toEqual(address);
        expect(events[index].newValue).toBe(index);
      });

      unsubscribe();
    });

    test('should handle formula evaluation in events', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push(event);
      });

      // Set base values
      const addressA1: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const addressB1: SimpleCellAddress = { sheet: sheetId, col: 1, row: 0 };
      const addressC1: SimpleCellAddress = { sheet: sheetId, col: 2, row: 0 };

      engine.setCellContent(addressA1, 10);
      engine.setCellContent(addressB1, 20);
      engine.setCellContent(addressC1, '=A1+B1');

      expect(events).toHaveLength(3);
      expect(events[2]).toEqual({
        address: addressC1,
        oldValue: undefined,
        newValue: 30
      });

      unsubscribe();
    });
  });

  describe('event timing', () => {
    test('should emit events synchronously', () => {
      let eventFired = false;
      const unsubscribe = engine.on('cell-changed', () => {
        eventFired = true;
      });

      const address: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContent(address, 42);

      // Event should have fired synchronously
      expect(eventFired).toBe(true);

      unsubscribe();
    });

    test('should emit events in correct order for dependent cells', () => {
      const events: any[] = [];
      const unsubscribe = engine.on('cell-changed', (event) => {
        events.push({
          address: `${String.fromCharCode(65 + event.address.col)}${event.address.row + 1}`,
          value: event.newValue
        });
      });

      // Set up dependencies: C1 = A1 + B1
      const addressA1: SimpleCellAddress = { sheet: sheetId, col: 0, row: 0 };
      const addressB1: SimpleCellAddress = { sheet: sheetId, col: 1, row: 0 };
      const addressC1: SimpleCellAddress = { sheet: sheetId, col: 2, row: 0 };

      engine.setCellContent(addressA1, 10);
      engine.setCellContent(addressB1, 20);
      engine.setCellContent(addressC1, '=A1+B1');

      // Clear events and test dependency update
      events.length = 0;
      
      // Update A1, should trigger C1 recalculation
      engine.setCellContent(addressA1, 15);

      // Should have exactly 2 events: A1 change and C1 recalculation
      expect(events).toHaveLength(2);
      expect(events[0]).toEqual({ address: 'A1', value: 15 });
      expect(events[1]).toEqual({ address: 'C1', value: 35 });

      unsubscribe();
    });
  });
});