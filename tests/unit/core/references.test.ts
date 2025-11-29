import { describe, it, expect, beforeEach } from 'bun:test';
import { FormulaEngine } from '../../../src/core/engine';
import type { RangeAddress } from '../../../src/core/types';

describe('Reference Tracking', () => {
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook('wb1');
    engine.addSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });
  });

  describe('createRef and getRefAddress', () => {
    it('should create reference and return UUID', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 4 }, row: { type: 'number', value: 9 } },
        },
      };

      const refId = engine.createRef(address);

      expect(refId).toBeDefined();
      expect(typeof refId).toBe('string');
      expect(refId.length).toBeGreaterThan(0);
    });

    it('should retrieve address by UUID', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 2 }, row: { type: 'number', value: 3 } },
        },
      };

      const refId = engine.createRef(address);
      const retrieved = engine.getRefAddress(refId);

      expect(retrieved).toEqual(address);
    });

    it('should return undefined for non-existent UUID', () => {
      const retrieved = engine.getRefAddress('non-existent-uuid');
      expect(retrieved).toBeUndefined();
    });

    it('should create multiple independent references', () => {
      const addr1: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const addr2: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 5, row: 5 },
          end: { col: { type: 'number', value: 10 }, row: { type: 'number', value: 10 } },
        },
      };

      const ref1 = engine.createRef(addr1);
      const ref2 = engine.createRef(addr2);

      expect(ref1).not.toBe(ref2);
      expect(engine.getRefAddress(ref1)).toEqual(addr1);
      expect(engine.getRefAddress(ref2)).toEqual(addr2);
    });
  });

  describe('deleteRef', () => {
    it('should delete reference', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId = engine.createRef(address);
      expect(engine.getRefAddress(refId)).toEqual(address);

      const deleted = engine.deleteRef(refId);
      expect(deleted).toBe(true);
      expect(engine.getRefAddress(refId)).toBeUndefined();
    });

    it('should return false for non-existent UUID', () => {
      const deleted = engine.deleteRef('non-existent-uuid');
      expect(deleted).toBe(false);
    });
  });

  describe('Sheet rename updates', () => {
    it('should update reference when sheet is renamed', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 5 }, row: { type: 'number', value: 5 } },
        },
      };

      const refId = engine.createRef(address);

      engine.renameSheet({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        newSheetName: 'Sales',
      });

      const updated = engine.getRefAddress(refId);
      expect(updated?.sheetName).toBe('Sales');
      expect(updated?.workbookName).toBe('wb1');
      expect(updated?.range).toEqual(address.range);
    });

    it('should not update references to other sheets', () => {
      engine.addSheet({ workbookName: 'wb1', sheetName: 'Sheet2' });

      const ref1Address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const ref2Address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet2',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const ref1 = engine.createRef(ref1Address);
      const ref2 = engine.createRef(ref2Address);

      engine.renameSheet({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        newSheetName: 'Renamed',
      });

      expect(engine.getRefAddress(ref1)?.sheetName).toBe('Renamed');
      expect(engine.getRefAddress(ref2)?.sheetName).toBe('Sheet2');
    });

    it('should handle multiple references to same sheet', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const ref1 = engine.createRef(address);
      const ref2 = engine.createRef(address);

      engine.renameSheet({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        newSheetName: 'Data',
      });

      expect(engine.getRefAddress(ref1)?.sheetName).toBe('Data');
      expect(engine.getRefAddress(ref2)?.sheetName).toBe('Data');
    });
  });

  describe('Workbook rename updates', () => {
    it('should update reference when workbook is renamed', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 3 }, row: { type: 'number', value: 3 } },
        },
      };

      const refId = engine.createRef(address);

      engine.renameWorkbook({
        workbookName: 'wb1',
        newWorkbookName: 'MyWorkbook',
      });

      const updated = engine.getRefAddress(refId);
      expect(updated?.workbookName).toBe('MyWorkbook');
      expect(updated?.sheetName).toBe('Sheet1');
      expect(updated?.range).toEqual(address.range);
    });

    it('should update all sheets in renamed workbook', () => {
      engine.addSheet({ workbookName: 'wb1', sheetName: 'Sheet2' });

      const ref1: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const ref2: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet2',
        range: {
          start: { col: 1, row: 1 },
          end: { col: { type: 'number', value: 1 }, row: { type: 'number', value: 1 } },
        },
      };

      const refId1 = engine.createRef(ref1);
      const refId2 = engine.createRef(ref2);

      engine.renameWorkbook({
        workbookName: 'wb1',
        newWorkbookName: 'NewWB',
      });

      expect(engine.getRefAddress(refId1)?.workbookName).toBe('NewWB');
      expect(engine.getRefAddress(refId2)?.workbookName).toBe('NewWB');
    });

    it('should not update references to other workbooks', () => {
      engine.addWorkbook('wb2');
      engine.addSheet({ workbookName: 'wb2', sheetName: 'Sheet1' });

      const ref1: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const ref2: RangeAddress = {
        workbookName: 'wb2',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId1 = engine.createRef(ref1);
      const refId2 = engine.createRef(ref2);

      engine.renameWorkbook({
        workbookName: 'wb1',
        newWorkbookName: 'Renamed',
      });

      expect(engine.getRefAddress(refId1)?.workbookName).toBe('Renamed');
      expect(engine.getRefAddress(refId2)?.workbookName).toBe('wb2');
    });
  });

  describe('Invalidation on delete', () => {
    it('should invalidate references when sheet is removed', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId = engine.createRef(address);
      expect(engine.getRefAddress(refId)).toEqual(address);

      engine.removeSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      expect(engine.getRefAddress(refId)).toBeUndefined();
    });

    it('should invalidate all references when workbook is removed', () => {
      engine.addSheet({ workbookName: 'wb1', sheetName: 'Sheet2' });

      const ref1: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const ref2: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet2',
        range: {
          start: { col: 1, row: 1 },
          end: { col: { type: 'number', value: 1 }, row: { type: 'number', value: 1 } },
        },
      };

      const refId1 = engine.createRef(ref1);
      const refId2 = engine.createRef(ref2);

      engine.removeWorkbook('wb1');

      expect(engine.getRefAddress(refId1)).toBeUndefined();
      expect(engine.getRefAddress(refId2)).toBeUndefined();
    });

    it('should return undefined for invalidated references', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 5 }, row: { type: 'number', value: 5 } },
        },
      };

      const refId = engine.createRef(address);
      
      // Remove the sheet
      engine.removeSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      // Reference should return undefined
      const retrieved = engine.getRefAddress(refId);
      expect(retrieved).toBeUndefined();
    });

    it('should track invalid references via getInvalidRefs', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId = engine.createRef(address);

      // Initially no invalid refs
      expect(engine.getInvalidRefs()).toEqual([]);

      // Remove sheet
      engine.removeSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      // Now should have one invalid ref
      const invalidRefs = engine.getInvalidRefs();
      expect(invalidRefs).toContain(refId);
      expect(invalidRefs.length).toBe(1);
    });
  });

  describe('Serialization', () => {
    it('should serialize and deserialize references', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 10 }, row: { type: 'number', value: 20 } },
        },
      };

      const refId = engine.createRef(address);

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty();
      newEngine.resetToSerializedEngine(serialized);

      const retrieved = newEngine.getRefAddress(refId);
      expect(retrieved).toEqual(address);
    });

    it('should preserve UUIDs across serialization', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const originalRefId = engine.createRef(address);

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty();
      newEngine.resetToSerializedEngine(serialized);

      // Same UUID should work in new engine
      const retrieved = newEngine.getRefAddress(originalRefId);
      expect(retrieved).toEqual(address);
    });

    it('should preserve validity state', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId = engine.createRef(address);
      engine.removeSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      // Reference is invalid
      expect(engine.getRefAddress(refId)).toBeUndefined();

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty();
      newEngine.resetToSerializedEngine(serialized);

      // Still invalid in new engine
      expect(newEngine.getRefAddress(refId)).toBeUndefined();
    });

    it('should preserve invalid references', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId = engine.createRef(address);
      engine.removeSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      expect(engine.getInvalidRefs()).toContain(refId);

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty();
      newEngine.resetToSerializedEngine(serialized);

      expect(newEngine.getInvalidRefs()).toContain(refId);
    });
  });

  describe('Edge cases', () => {
    it('should handle references to same cell as range', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 5, row: 5 },
          end: { col: { type: 'number', value: 5 }, row: { type: 'number', value: 5 } },
        },
      };

      const refId = engine.createRef(address);
      const retrieved = engine.getRefAddress(refId);

      expect(retrieved).toEqual(address);
    });

    it('should not clone references when cloning workbook', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      };

      const refId = engine.createRef(address);

      engine.cloneWorkbook('wb1', 'wb2');

      // Original reference still points to wb1
      const retrieved = engine.getRefAddress(refId);
      expect(retrieved?.workbookName).toBe('wb1');
      expect(retrieved?.sheetName).toBe('Sheet1');

      // If we rename the clone, original ref unchanged
      engine.renameWorkbook({ workbookName: 'wb2', newWorkbookName: 'wb3' });
      expect(engine.getRefAddress(refId)?.workbookName).toBe('wb1');
    });

    it('should handle infinite ranges in references', () => {
      const address: RangeAddress = {
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'infinity', sign: 'positive' }, row: { type: 'number', value: 10 } },
        },
      };

      const refId = engine.createRef(address);
      const retrieved = engine.getRefAddress(refId);

      expect(retrieved).toEqual(address);
    });
  });

  describe('Integration with metadata', () => {
    it('should support storing refIds in sheet metadata', () => {
      interface SheetMeta {
        textBoxes?: Array<{ id: string; anchorRef: string; content: string }>;
      }

      const typedEngine = FormulaEngine.buildEmpty<unknown, SheetMeta, unknown>();
      typedEngine.addWorkbook('wb1');
      typedEngine.addSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      const anchorRef = typedEngine.createRef({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 2 }, row: { type: 'number', value: 5 } },
        },
      });

      typedEngine.setSheetMetadata({ workbookName: 'wb1', sheetName: 'Sheet1' }, {
        textBoxes: [
          { id: 'tb1', anchorRef, content: 'Sales Chart' },
        ],
      });

      // Rename sheet
      typedEngine.renameSheet({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        newSheetName: 'Sales',
      });

      // Get updated address from stored refId
      const metadata = typedEngine.getSheetMetadata({ workbookName: 'wb1', sheetName: 'Sales' });
      const textBoxAnchor = metadata?.textBoxes?.[0]?.anchorRef;
      const anchorAddress = typedEngine.getRefAddress(textBoxAnchor!);

      expect(anchorAddress?.sheetName).toBe('Sales');
      expect(anchorAddress?.workbookName).toBe('wb1');
    });

    it('should handle cleanup of invalid refs from metadata', () => {
      interface SheetMeta {
        anchors?: string[];
      }

      const typedEngine = FormulaEngine.buildEmpty<unknown, SheetMeta, unknown>();
      typedEngine.addWorkbook('wb1');
      typedEngine.addSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      const ref1 = typedEngine.createRef({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: 'number', value: 0 }, row: { type: 'number', value: 0 } },
        },
      });

      const ref2 = typedEngine.createRef({
        workbookName: 'wb1',
        sheetName: 'Sheet1',
        range: {
          start: { col: 1, row: 1 },
          end: { col: { type: 'number', value: 1 }, row: { type: 'number', value: 1 } },
        },
      });

      typedEngine.setSheetMetadata({ workbookName: 'wb1', sheetName: 'Sheet1' }, {
        anchors: [ref1, ref2],
      });

      // Remove sheet
      typedEngine.removeSheet({ workbookName: 'wb1', sheetName: 'Sheet1' });

      // Check for invalid refs
      const invalidRefs = typedEngine.getInvalidRefs();
      expect(invalidRefs).toContain(ref1);
      expect(invalidRefs).toContain(ref2);

      // Clean them up
      invalidRefs.forEach((refId) => typedEngine.deleteRef(refId));

      expect(typedEngine.getInvalidRefs()).toEqual([]);
    });
  });
});

