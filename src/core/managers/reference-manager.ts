/**
 * ReferenceManager - Manages tracked references for external elements
 * 
 * Allows consumers to create stable references to ranges that automatically
 * update when sheets/workbooks are renamed and become invalid when deleted.
 */

import type { RangeAddress, TrackedReference } from "../types";

export class ReferenceManager {
  private references: Map<string, TrackedReference>;

  constructor() {
    this.references = new Map();
  }

  /**
   * Create a new tracked reference
   * Returns UUID for the reference
   */
  createRef(address: RangeAddress): string {
    const uuid = crypto.randomUUID();
    this.references.set(uuid, {
      id: uuid,
      address: {
        workbookName: address.workbookName,
        sheetName: address.sheetName,
        range: address.range,
      },
      isValid: true,
    });
    return uuid;
  }

  /**
   * Get the current address for a reference
   * Returns undefined if reference doesn't exist or is invalid
   */
  getRefAddress(refId: string): RangeAddress | undefined {
    const ref = this.references.get(refId);
    if (!ref || !ref.isValid) {
      return undefined;
    }
    return {
      workbookName: ref.address.workbookName,
      sheetName: ref.address.sheetName,
      range: ref.address.range,
    };
  }

  /**
   * Delete a reference
   * Returns true if reference was deleted, false if it didn't exist
   */
  deleteRef(refId: string): boolean {
    return this.references.delete(refId);
  }

  /**
   * Get all invalid reference IDs
   */
  getInvalidRefs(): string[] {
    const invalid: string[] = [];
    for (const [id, ref] of this.references) {
      if (!ref.isValid) {
        invalid.push(id);
      }
    }
    return invalid;
  }

  /**
   * Update references when sheet is renamed
   */
  updateSheetName(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    for (const ref of this.references.values()) {
      if (
        ref.address.workbookName === workbookName &&
        ref.address.sheetName === oldSheetName
      ) {
        ref.address.sheetName = newSheetName;
      }
    }
  }

  /**
   * Update references when workbook is renamed
   */
  updateWorkbookName(oldWorkbookName: string, newWorkbookName: string): void {
    for (const ref of this.references.values()) {
      if (ref.address.workbookName === oldWorkbookName) {
        ref.address.workbookName = newWorkbookName;
      }
    }
  }

  /**
   * Mark references as invalid when sheet is removed
   */
  invalidateSheet(workbookName: string, sheetName: string): void {
    for (const ref of this.references.values()) {
      if (
        ref.address.workbookName === workbookName &&
        ref.address.sheetName === sheetName
      ) {
        ref.isValid = false;
      }
    }
  }

  /**
   * Mark references as invalid when workbook is removed
   */
  invalidateWorkbook(workbookName: string): void {
    for (const ref of this.references.values()) {
      if (ref.address.workbookName === workbookName) {
        ref.isValid = false;
      }
    }
  }

  /**
   * Get all references for serialization
   */
  getAllReferences(): Map<string, TrackedReference> {
    return new Map(this.references);
  }

  /**
   * Restore references from serialization
   */
  resetReferences(refs: Map<string, TrackedReference>): void {
    this.references.clear();
    for (const [id, ref] of refs) {
      this.references.set(id, {
        id: ref.id,
        address: {
          workbookName: ref.address.workbookName,
          sheetName: ref.address.sheetName,
          range: ref.address.range,
        },
        isValid: ref.isValid,
      });
    }
  }
}

