/**
 * The common shape shared by all reversible engine mutations.
 *
 * Concrete history steps are expected to form a discriminated union by
 * extending this interface and narrowing `kind`. The manager deliberately
 * does not know how to apply a step; it only owns retention and accounting.
 */
export interface HistoryStep {
  readonly kind: string;

  /**
   * Approximate retained size of this step in bytes. Supplying this avoids an
   * estimator walk when the mutation already knows the size of its payload.
   */
  readonly estimatedBytes?: number;
}

/** One user-visible undo/redo operation, applied in step order. */
export interface HistoryEntry<TStep extends HistoryStep = HistoryStep> {
  readonly steps: readonly TStep[];
  readonly estimatedBytes: number;
  readonly label?: string;
}

export interface CreateHistoryEntryOptions {
  readonly label?: string;

  /**
   * Overrides per-step accounting. This is useful for payloads whose retained
   * size can be calculated while the change is captured.
   */
  readonly estimatedBytes?: number;
}

const REFERENCE_BYTES = 8;
const OBJECT_HEADER_BYTES = 16;
const ARRAY_HEADER_BYTES = 24;
const MAP_OR_SET_HEADER_BYTES = 32;
const MAX_SAFE_HISTORY_ARRAY_LENGTH = 100_000;
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
const INTRINSIC_ARRAY_BUFFER_VIEW_PROTOTYPES = new Set<object>([
  DataView.prototype,
  Int8Array.prototype,
  Uint8Array.prototype,
  Uint8ClampedArray.prototype,
  Int16Array.prototype,
  Uint16Array.prototype,
  Int32Array.prototype,
  Uint32Array.prototype,
  Float32Array.prototype,
  Float64Array.prototype,
  BigInt64Array.prototype,
  BigUint64Array.prototype,
]);
const ARRAY_BUFFER_BYTE_LENGTH_GETTER = Object.getOwnPropertyDescriptor(
  ArrayBuffer.prototype,
  "byteLength"
)!.get as (this: ArrayBuffer) => number;
const REGEXP_SOURCE_GETTER = Object.getOwnPropertyDescriptor(
  RegExp.prototype,
  "source"
)!.get as (this: RegExp) => string;
const DATE_GET_TIME = Date.prototype.getTime;

function propertyKeyBytes(key: PropertyKey): number {
  if (typeof key === "string") {
    return key.length * 2;
  }
  if (typeof key === "symbol") {
    return (key.description?.length ?? 0) * 2 + REFERENCE_BYTES;
  }
  return String(key).length * 2;
}

/**
 * Returns whether history can detach and account for every reference reachable
 * from a value. Unsupported opaque values become an undo barrier rather than
 * silently retaining an unbounded graph behind a tiny estimate.
 */
export function isHistoryValueSafelyRetainable(value: unknown): boolean {
  const seen = new WeakSet<object>();

  const safe = (current: unknown, depth = 0): boolean => {
    if (depth > 1_024) {
      return false;
    }
    if (current === null || current === undefined) {
      return true;
    }
    if (typeof current === "function") {
      return false;
    }
    if (typeof current === "bigint") {
      // JavaScript has no allocation-free bit-length query. Converting an
      // attacker-sized bigint to a string just to budget it can itself OOM.
      return false;
    }
    if (typeof current !== "object") {
      return true;
    }
    if (seen.has(current)) {
      return true;
    }
    seen.add(current);

    try {
      const prototype = Object.getPrototypeOf(current);
      if (current instanceof Date) {
        if (
          prototype !== Date.prototype ||
          Reflect.ownKeys(current).length !== 0
        ) {
          return false;
        }
        DATE_GET_TIME.call(current);
        return true;
      }
      if (current instanceof ArrayBuffer) {
        if (
          prototype !== ArrayBuffer.prototype ||
          Reflect.ownKeys(current).length !== 0
        ) {
          return false;
        }
        ARRAY_BUFFER_BYTE_LENGTH_GETTER.call(current);
        return true;
      }
      if (ArrayBuffer.isView(current)) {
        // Indexed typed-array keys are proportional to byte length and must
        // never be enumerated here. The intrinsic buffer is cloned/accounted
        // by byteLength in O(1).
        return (
          INTRINSIC_ARRAY_BUFFER_VIEW_PROTOTYPES.has(prototype) &&
          !["buffer", "byteLength", "byteOffset", "length", "constructor"].some(
            (key) => Object.prototype.hasOwnProperty.call(current, key)
          )
        );
      }
      if (current instanceof RegExp) {
        if (
          prototype !== RegExp.prototype ||
          !Reflect.ownKeys(current).every((key) => key === "lastIndex")
        ) {
          return false;
        }
        REGEXP_SOURCE_GETTER.call(current);
        return true;
      }
      if (current instanceof Map) {
        if (
          prototype !== Map.prototype ||
          Reflect.ownKeys(current).length > 0
        ) {
          return false;
        }
        for (const [key, mapValue] of current) {
          if (
            (typeof key === "object" && key !== null) ||
            typeof key === "function"
          ) {
            return false;
          }
          if (!safe(key, depth + 1) || !safe(mapValue, depth + 1)) {
            return false;
          }
        }
        return true;
      }
      if (current instanceof Set) {
        if (
          prototype !== Set.prototype ||
          Reflect.ownKeys(current).length > 0
        ) {
          return false;
        }
        for (const item of current) {
          if (
            (typeof item === "object" && item !== null) ||
            typeof item === "function"
          ) {
            return false;
          }
          if (!safe(item, depth + 1)) {
            return false;
          }
        }
        return true;
      }
      if (
        Array.isArray(current) &&
        (prototype !== Array.prototype ||
          current.length > MAX_SAFE_HISTORY_ARRAY_LENGTH)
      ) {
        return false;
      }

      if (
        !Array.isArray(current) &&
        (!(current instanceof Error) ||
          !INTRINSIC_ERROR_PROTOTYPES.has(prototype)) &&
        prototype !== Object.prototype &&
        prototype !== null
      ) {
        return false;
      }

      for (const key of Reflect.ownKeys(current)) {
        const descriptor = Object.getOwnPropertyDescriptor(current, key);
        if (!descriptor) {
          return false;
        }
        if (!("value" in descriptor)) {
          return false;
        }
        if (!safe(descriptor.value, depth + 1)) {
          return false;
        }
      }
      return true;
    } catch {
      return false;
    }
  };

  return safe(value);
}

function assertValidByteSize(value: number, name: string): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative safe integer`);
  }
}

/**
 * Estimates the retained size of a history payload without serializing it.
 *
 * This intentionally favors a predictable, allocation-light approximation
 * over VM-specific object sizing. Shared or cyclic objects are counted once.
 * Callers on hot paths can provide `estimatedBytes` directly instead.
 */
export function estimateHistoryValueBytes(value: unknown): number {
  const seen = new WeakSet<object>();

  const estimate = (current: unknown): number => {
    if (current === null || current === undefined) {
      return REFERENCE_BYTES;
    }

    switch (typeof current) {
      case "boolean":
      case "number":
        return 8;
      case "bigint":
        return Number.MAX_SAFE_INTEGER;
      case "string":
        return current.length * 2;
      case "symbol":
        return REFERENCE_BYTES + (current.description?.length ?? 0) * 2;
      case "function":
        return REFERENCE_BYTES;
      case "object":
        break;
    }

    if (seen.has(current)) {
      return REFERENCE_BYTES;
    }
    seen.add(current);

    if (current instanceof ArrayBuffer) {
      try {
        return (
          OBJECT_HEADER_BYTES + ARRAY_BUFFER_BYTE_LENGTH_GETTER.call(current)
        );
      } catch {
        return Number.MAX_SAFE_INTEGER;
      }
    }

    if (ArrayBuffer.isView(current)) {
      try {
        return OBJECT_HEADER_BYTES + current.byteLength;
      } catch {
        return Number.MAX_SAFE_INTEGER;
      }
    }

    if (current instanceof Date) {
      return OBJECT_HEADER_BYTES + 8;
    }

    if (Array.isArray(current)) {
      if (current.length > MAX_SAFE_HISTORY_ARRAY_LENGTH) {
        return Number.MAX_SAFE_INTEGER;
      }
      let bytes = ARRAY_HEADER_BYTES + current.length * REFERENCE_BYTES;
      for (const key of Reflect.ownKeys(current)) {
        if (key === "length") {
          continue;
        }
        bytes += REFERENCE_BYTES + propertyKeyBytes(key);
        const descriptor = Object.getOwnPropertyDescriptor(current, key);
        if (!descriptor) {
          continue;
        }
        if ("value" in descriptor) {
          bytes += estimate(descriptor.value);
        } else {
          bytes +=
            (descriptor.get ? REFERENCE_BYTES : 0) +
            (descriptor.set ? REFERENCE_BYTES : 0);
        }
      }
      return bytes;
    }

    if (current instanceof Map) {
      let bytes = MAP_OR_SET_HEADER_BYTES + current.size * REFERENCE_BYTES * 2;
      for (const [key, mapValue] of current) {
        bytes += estimate(key) + estimate(mapValue);
      }
      return bytes;
    }

    if (current instanceof Set) {
      let bytes = MAP_OR_SET_HEADER_BYTES + current.size * REFERENCE_BYTES;
      for (const item of current) {
        bytes += estimate(item);
      }
      return bytes;
    }

    if (current instanceof RegExp) {
      try {
        return (
          OBJECT_HEADER_BYTES +
          REGEXP_SOURCE_GETTER.call(current).length * 2 +
          current.flags.length * 2 +
          8
        );
      } catch {
        return Number.MAX_SAFE_INTEGER;
      }
    }

    let bytes = OBJECT_HEADER_BYTES;
    let keys: PropertyKey[];
    try {
      keys = Reflect.ownKeys(current);
    } catch {
      return Number.MAX_SAFE_INTEGER;
    }
    for (const key of keys) {
      bytes += REFERENCE_BYTES + propertyKeyBytes(key);
      let descriptor: PropertyDescriptor | undefined;
      try {
        descriptor = Object.getOwnPropertyDescriptor(current, key);
      } catch {
        return Number.MAX_SAFE_INTEGER;
      }
      if (!descriptor) {
        continue;
      }
      if ("value" in descriptor) {
        bytes += estimate(descriptor.value);
      } else {
        bytes +=
          (descriptor.get ? REFERENCE_BYTES : 0) +
          (descriptor.set ? REFERENCE_BYTES : 0);
      }
    }
    return bytes;
  };

  return estimate(value);
}

export function estimateHistoryStepsBytes<TStep extends HistoryStep>(
  steps: readonly TStep[]
): number {
  let bytes = 0;
  for (const step of steps) {
    if (step.estimatedBytes !== undefined) {
      assertValidByteSize(step.estimatedBytes, "step.estimatedBytes");
      bytes += step.estimatedBytes;
    } else {
      bytes += estimateHistoryValueBytes(step);
    }

    if (!Number.isSafeInteger(bytes)) {
      throw new Error(
        "history entry estimated byte size exceeds safe integer range"
      );
    }
  }
  return bytes;
}

/**
 * Builds an immutable entry envelope. Step payloads themselves are not cloned;
 * mutation capture code must not mutate data after handing it to history.
 */
export function createHistoryEntry<TStep extends HistoryStep>(
  steps: readonly TStep[],
  options: CreateHistoryEntryOptions = {}
): HistoryEntry<TStep> {
  const estimatedBytes =
    options.estimatedBytes ?? estimateHistoryStepsBytes(steps);
  assertValidByteSize(estimatedBytes, "history entry estimatedBytes");

  return {
    steps: [...steps],
    estimatedBytes,
    ...(options.label === undefined ? {} : { label: options.label }),
  };
}
