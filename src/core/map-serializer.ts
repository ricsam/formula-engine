// Replacer: used with JSON.stringify
export function replacer(key: string, value: any): any {
  if (value instanceof Set) {
    return {
      __type: "Set",
      values: Array.from(value),
    };
  }
  if (value instanceof Map) {
    return {
      __type: "Map",
      entries: Array.from(value.entries()),
    };
  }
  return value;
}

// Reviver: used with JSON.parse
export function reviver(key: string, value: any): any {
  if (value && value.__type === "Set") {
    return new Set(value.values);
  }
  if (value && value.__type === "Map") {
    return new Map(value.entries);
  }
  return value;
}

export const serialize = (data: unknown): string => {
  return JSON.stringify(data, replacer);
};

export const deserialize = (data: string): unknown => {
  return JSON.parse(data, reviver);
};
