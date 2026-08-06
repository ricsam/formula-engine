import type { NamedExpressionMutation } from "./managers/named-expression-manager";
import type { RangeMetadataDataChange } from "./managers/range-metadata-manager";
import type { ReferenceDataChange } from "./managers/reference-manager";
import type { StyleDataChange } from "./managers/style-manager";
import type { TableMutation } from "./managers/table-manager";
import type { WorkbookDataChange } from "./managers/workbook-manager";
import { cloneMutationValue } from "./managers/mutation-observer";
import type { Sheet, Workbook } from "./types";

export type WorkbookScopeState = {
  workbookName: string;
  workbookOrder: string[];
  workbook?: Workbook;
};

export type SheetScopeState = {
  workbookName: string;
  sheetName: string;
  sheetOrder: string[];
  sheet?: Sheet;
};

export type EngineHistoryStep<TRangeMetadata = unknown> =
  | {
      kind: "workbook-data";
      patches: WorkbookDataChange[][];
      atomicGroupId?: number;
      committed?: boolean;
      appendOnlyCellContent?: true;
      sequentialCellContentDeletions?: true;
      estimatedBytes?: number;
    }
  | {
      kind: "named-expression-data";
      changes: NamedExpressionMutation[];
      estimatedBytes?: number;
    }
  | {
      kind: "table-data";
      changes: TableMutation[];
      estimatedBytes?: number;
    }
  | {
      kind: "style-data";
      changes: StyleDataChange[];
      estimatedBytes?: number;
    }
  | {
      kind: "range-metadata-data";
      changes: RangeMetadataDataChange<TRangeMetadata>[];
      estimatedBytes?: number;
    }
  | {
      kind: "reference-data";
      changes: ReferenceDataChange[];
      estimatedBytes?: number;
    }
  | {
      kind: "rename-sheet";
      workbookName: string;
      before: string;
      after: string;
      estimatedBytes?: number;
    }
  | {
      kind: "rename-workbook";
      before: string;
      after: string;
      estimatedBytes?: number;
    }
  | {
      kind: "workbook-scope";
      before: WorkbookScopeState;
      after: WorkbookScopeState;
      estimatedBytes?: number;
    }
  | {
      kind: "sheet-scope";
      before: SheetScopeState;
      after: SheetScopeState;
      estimatedBytes?: number;
    };

/**
 * History owns detached mutation payloads. Serialization is intentionally
 * limited to the affected manager/scope rather than the complete engine.
 */
export function cloneHistoryValue<T>(value: T): T {
  return cloneMutationValue(value);
}

export function historyValuesEqual(left: unknown, right: unknown): boolean {
  const visited = new WeakMap<object, WeakSet<object>>();

  const equal = (leftValue: unknown, rightValue: unknown): boolean => {
    if (Object.is(leftValue, rightValue)) {
      return true;
    }
    if (
      leftValue === null ||
      rightValue === null ||
      typeof leftValue !== "object" ||
      typeof rightValue !== "object"
    ) {
      return false;
    }

    let rightValues = visited.get(leftValue);
    if (rightValues?.has(rightValue)) {
      return true;
    }
    if (!rightValues) {
      rightValues = new WeakSet<object>();
      visited.set(leftValue, rightValues);
    }
    rightValues.add(rightValue);

    if (
      Object.getPrototypeOf(leftValue) !== Object.getPrototypeOf(rightValue)
    ) {
      return false;
    }

    if (leftValue instanceof Date || rightValue instanceof Date) {
      return (
        leftValue instanceof Date &&
        rightValue instanceof Date &&
        leftValue.getTime() === rightValue.getTime()
      );
    }

    if (leftValue instanceof RegExp || rightValue instanceof RegExp) {
      return (
        leftValue instanceof RegExp &&
        rightValue instanceof RegExp &&
        leftValue.source === rightValue.source &&
        leftValue.flags === rightValue.flags &&
        leftValue.lastIndex === rightValue.lastIndex
      );
    }

    if (leftValue instanceof Map || rightValue instanceof Map) {
      if (
        !(leftValue instanceof Map) ||
        !(rightValue instanceof Map) ||
        leftValue.size !== rightValue.size
      ) {
        return false;
      }
      const rightEntries = rightValue.entries();
      for (const [leftKey, leftMapValue] of leftValue) {
        const rightEntry = rightEntries.next();
        if (
          rightEntry.done ||
          !Object.is(leftKey, rightEntry.value[0]) ||
          !equal(leftMapValue, rightEntry.value[1])
        ) {
          return false;
        }
      }
      return true;
    }

    if (leftValue instanceof Set || rightValue instanceof Set) {
      if (
        !(leftValue instanceof Set) ||
        !(rightValue instanceof Set) ||
        leftValue.size !== rightValue.size
      ) {
        return false;
      }
      const rightValues = rightValue.values();
      for (const leftSetValue of leftValue) {
        const rightValueResult = rightValues.next();
        if (
          rightValueResult.done ||
          !Object.is(leftSetValue, rightValueResult.value)
        ) {
          return false;
        }
      }
      return true;
    }

    if (leftValue instanceof ArrayBuffer || rightValue instanceof ArrayBuffer) {
      if (
        !(leftValue instanceof ArrayBuffer) ||
        !(rightValue instanceof ArrayBuffer) ||
        leftValue.byteLength !== rightValue.byteLength
      ) {
        return false;
      }
      const leftBytes = new Uint8Array(leftValue);
      const rightBytes = new Uint8Array(rightValue);
      return leftBytes.every((byte, index) => byte === rightBytes[index]);
    }

    if (ArrayBuffer.isView(leftValue) || ArrayBuffer.isView(rightValue)) {
      if (!ArrayBuffer.isView(leftValue) || !ArrayBuffer.isView(rightValue)) {
        return false;
      }
      if (leftValue.constructor !== rightValue.constructor) {
        return false;
      }
      const leftBytes = new Uint8Array(
        leftValue.buffer,
        leftValue.byteOffset,
        leftValue.byteLength
      );
      const rightBytes = new Uint8Array(
        rightValue.buffer,
        rightValue.byteOffset,
        rightValue.byteLength
      );
      if (leftBytes.length !== rightBytes.length) {
        return false;
      }
      return leftBytes.every((byte, index) => byte === rightBytes[index]);
    }

    if (Array.isArray(leftValue) !== Array.isArray(rightValue)) {
      return false;
    }

    const prototype = Object.getPrototypeOf(leftValue);
    if (
      !Array.isArray(leftValue) &&
      !(leftValue instanceof Error) &&
      prototype !== Object.prototype &&
      prototype !== null
    ) {
      return false;
    }

    const leftKeys = Reflect.ownKeys(leftValue);
    const rightKeys = Reflect.ownKeys(rightValue);
    if (leftKeys.length !== rightKeys.length) {
      return false;
    }
    return leftKeys.every((key) => {
      const leftDescriptor = Object.getOwnPropertyDescriptor(leftValue, key);
      const rightDescriptor = Object.getOwnPropertyDescriptor(rightValue, key);
      if (!leftDescriptor || !rightDescriptor) {
        return false;
      }
      if (
        leftDescriptor.configurable !== rightDescriptor.configurable ||
        leftDescriptor.enumerable !== rightDescriptor.enumerable
      ) {
        return false;
      }
      if ("value" in leftDescriptor || "value" in rightDescriptor) {
        return (
          "value" in leftDescriptor &&
          "value" in rightDescriptor &&
          leftDescriptor.writable === rightDescriptor.writable &&
          equal(leftDescriptor.value, rightDescriptor.value)
        );
      }
      return (
        Object.is(leftDescriptor.get, rightDescriptor.get) &&
        Object.is(leftDescriptor.set, rightDescriptor.set)
      );
    });
  };

  return equal(left, right);
}
