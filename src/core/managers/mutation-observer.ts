/** A value and its exact position in an ordered manager collection. */
export interface IndexedMutationValue<T> {
  readonly index: number;
  readonly value: T;
}

export type MutationDirection = "undo" | "redo";

export interface IndexedBeforeAfter<T> {
  readonly before?: IndexedMutationValue<T>;
  readonly after?: IndexedMutationValue<T>;
}

const INTRINSIC_ERROR_PROTOTYPES = new Set<object>([
  Error.prototype,
  EvalError.prototype,
  RangeError.prototype,
  ReferenceError.prototype,
  SyntaxError.prototype,
  TypeError.prototype,
  URIError.prototype,
  AggregateError.prototype,
]);

/**
 * Small reusable dispatcher for manager mutation observers.
 *
 * Managers build detached deltas only when `observed` is true. Nested batches
 * delay delivery, but each reported patch keeps its own index coordinate
 * system. Flattening sequential patches would make indexed replay ambiguous.
 */
export class MutationObserverDispatcher<TChange> {
  private batchDepth = 0;
  private suppressionDepth = 0;
  private pendingPatches: TChange[][] = [];

  constructor(
    private readonly observer?: (changes: readonly TChange[]) => void,
    private readonly shouldObserve: () => boolean = () => true,
    private readonly onSuppressedMutation?: () => void,
    private readonly detachValues = true
  ) {}

  retain<TValue>(value: TValue): TValue {
    return this.detachValues ? cloneMutationValue(value) : value;
  }

  get observed(): boolean {
    const active = this.observer !== undefined && this.shouldObserve();
    if (active && this.suppressionDepth > 0) {
      this.onSuppressedMutation?.();
      return false;
    }
    return active;
  }

  suppress<TResult>(callback: () => TResult): TResult {
    this.suppressionDepth++;
    try {
      return callback();
    } finally {
      this.suppressionDepth--;
    }
  }

  batch<TResult>(callback: () => TResult): TResult {
    this.batchDepth++;
    try {
      return callback();
    } finally {
      this.batchDepth--;
      if (this.batchDepth === 0 && this.pendingPatches.length > 0) {
        const patches = this.pendingPatches;
        this.pendingPatches = [];
        for (const patch of patches) {
          this.notify(patch);
        }
      }
    }
  }

  report(changes: readonly TChange[]): void {
    if (!this.observed || changes.length === 0) {
      return;
    }

    if (this.batchDepth > 0) {
      this.pendingPatches.push([...changes]);
      return;
    }

    this.notify(changes);
  }

  private notify(changes: readonly TChange[]): void {
    if (!this.observed) {
      return;
    }
    this.observer?.(changes);
  }
}

/** Clone retained observer data so later manager/caller writes cannot alter it. */
export function cloneMutationValue<T>(value: T): T {
  try {
    return cloneBestEffort(value, new WeakMap<object, unknown>());
  } catch {
    // Proxies and opaque host objects can reject reflection. Keep the original
    // valid object instead of manufacturing an invalid branded lookalike.
    return value;
  }
}

function cloneBestEffort<T>(value: T, seen: WeakMap<object, unknown>): T {
  if (
    value === null ||
    (typeof value !== "object" && typeof value !== "function")
  ) {
    return value;
  }
  if (typeof value === "function") {
    return value;
  }

  if (seen.has(value)) {
    return seen.get(value) as T;
  }

  if (value instanceof Date) {
    return new Date(value.getTime()) as T;
  }
  if (value instanceof RegExp) {
    const result = new RegExp(value.source, value.flags);
    result.lastIndex = value.lastIndex;
    return result as T;
  }
  if (value instanceof Error) {
    const prototype = Object.getPrototypeOf(value);
    const result = Object.create(
      INTRINSIC_ERROR_PROTOTYPES.has(prototype) ? prototype : Error.prototype
    ) as Error;
    seen.set(value, result);
    cloneOwnPropertyDescriptors(value, result, seen);
    return result as T;
  }
  if (value instanceof Map) {
    const result = new Map();
    seen.set(value, result);
    for (const [key, mapValue] of value) {
      result.set(cloneBestEffort(key, seen), cloneBestEffort(mapValue, seen));
    }
    return result as T;
  }
  if (value instanceof Set) {
    const result = new Set();
    seen.set(value, result);
    for (const item of value) {
      result.add(cloneBestEffort(item, seen));
    }
    return result as T;
  }
  if (value instanceof ArrayBuffer) {
    const result = new ArrayBuffer(value.byteLength);
    new Uint8Array(result).set(new Uint8Array(value));
    return result as T;
  }
  if (ArrayBuffer.isView(value)) {
    return cloneArrayBufferView(value) as T;
  }
  if (Array.isArray(value)) {
    const result: unknown[] = new Array(value.length);
    seen.set(value, result);
    cloneOwnPropertyDescriptors(value, result, seen, new Set(["length"]));
    return result as T;
  }

  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    return value;
  }

  const result = Object.create(Object.getPrototypeOf(value)) as Record<
    PropertyKey,
    unknown
  >;
  seen.set(value, result);
  cloneOwnPropertyDescriptors(value, result, seen);
  return result as T;
}

function cloneArrayBufferView(value: ArrayBufferView): ArrayBufferView {
  const buffer = new ArrayBuffer(value.byteLength);
  new Uint8Array(buffer).set(
    new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
  );

  if (value instanceof DataView) {
    return new DataView(buffer);
  }
  if (value instanceof Int8Array) {
    return new Int8Array(buffer);
  }
  if (value instanceof Uint8ClampedArray) {
    return new Uint8ClampedArray(buffer);
  }
  if (value instanceof Uint8Array) {
    return new Uint8Array(buffer);
  }
  if (value instanceof Int16Array) {
    return new Int16Array(buffer);
  }
  if (value instanceof Uint16Array) {
    return new Uint16Array(buffer);
  }
  if (value instanceof Int32Array) {
    return new Int32Array(buffer);
  }
  if (value instanceof Uint32Array) {
    return new Uint32Array(buffer);
  }
  if (value instanceof Float32Array) {
    return new Float32Array(buffer);
  }
  if (value instanceof Float64Array) {
    return new Float64Array(buffer);
  }
  if (value instanceof BigInt64Array) {
    return new BigInt64Array(buffer);
  }
  if (value instanceof BigUint64Array) {
    return new BigUint64Array(buffer);
  }

  // Unknown future view types are normalized to bytes rather than invoking a
  // caller-controlled constructor or retaining a custom prototype graph.
  return new Uint8Array(buffer);
}

function cloneOwnPropertyDescriptors(
  source: object,
  target: object,
  seen: WeakMap<object, unknown>,
  skippedKeys: ReadonlySet<PropertyKey> = new Set()
): void {
  for (const key of Reflect.ownKeys(source)) {
    if (skippedKeys.has(key)) {
      continue;
    }
    const descriptor = Object.getOwnPropertyDescriptor(source, key);
    if (!descriptor) {
      continue;
    }
    if ("value" in descriptor) {
      descriptor.value = cloneBestEffort(descriptor.value, seen);
    }
    Object.defineProperty(target, key, descriptor);
  }
}

/** Applies an ordered sparse patch without notifying a manager observer. */
export function applyIndexedChanges<
  TValue,
  TChange extends IndexedBeforeAfter<TValue>
>(
  current: readonly TValue[],
  changes: readonly TChange[],
  direction: MutationDirection
): TValue[] {
  const sourceKey = direction === "redo" ? "before" : "after";
  const targetKey = direction === "redo" ? "after" : "before";
  const removals = new Set(
    changes.flatMap((change) => {
      const source = change[sourceKey];
      return source ? [source.index] : [];
    })
  );
  const base = current.filter((_, index) => !removals.has(index));
  const insertions = changes
    .flatMap((change, insertionOrder) => {
      const target = change[targetKey];
      return target ? [{ ...target, insertionOrder }] : [];
    })
    .sort(
      (left, right) =>
        left.index - right.index || left.insertionOrder - right.insertionOrder
    );

  const result: TValue[] = [];
  let baseIndex = 0;
  let insertionIndex = 0;
  const finalLength = base.length + insertions.length;
  while (result.length < finalLength) {
    const insertion = insertions[insertionIndex];
    if (
      insertion &&
      Math.max(0, Math.min(insertion.index, finalLength - 1)) <= result.length
    ) {
      result.push(cloneMutationValue(insertion.value));
      insertionIndex++;
      continue;
    }

    const baseValue = base[baseIndex];
    if (baseIndex < base.length) {
      result.push(baseValue!);
      baseIndex++;
      continue;
    }

    if (insertion) {
      result.push(cloneMutationValue(insertion.value));
      insertionIndex++;
    }
  }

  return result;
}
