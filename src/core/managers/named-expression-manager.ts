import type { NamedExpression } from "../types";
import type { NamedExpressionManagerSnapshot } from "../engine-snapshot";
import { renameNamedExpressionInFormula } from "../named-expression-renamer";
import type { EventManager } from "./event-manager";
import type { NamedExpressionNode } from "../../parser/ast";
import type { EvaluationContext } from "../../evaluator/evaluation-context";
import { getNamedExpressionResourceKey } from "../resource-keys";
import type { MutationDirection } from "./mutation-observer";

export type NamedExpressionScope =
  | { type: "global" }
  | { type: "workbook"; workbookName: string }
  | { type: "sheet"; workbookName: string; sheetName: string };

export type NamedExpressionBucket =
  | { type: "workbook"; workbookName: string }
  | { type: "sheet-workbook"; workbookName: string }
  | { type: "sheet"; workbookName: string; sheetName: string };

export type NamedExpressionEntryState = {
  scope: NamedExpressionScope;
  expressionName: string;
  expression: NamedExpression;
  index: number;
};

export type NamedExpressionBucketState = {
  bucket: NamedExpressionBucket;
  index: number;
};

export type NamedExpressionMutation =
  | {
      kind: "named-expression";
      before?: NamedExpressionEntryState;
      after?: NamedExpressionEntryState;
    }
  | {
      kind: "named-expression-bucket";
      before?: NamedExpressionBucketState;
      after?: NamedExpressionBucketState;
    };

export type NamedExpressionMutationObserver = (
  changes: readonly NamedExpressionMutation[]
) => void;

type NamedExpressionState = {
  entries: NamedExpressionEntryState[];
  buckets: NamedExpressionBucketState[];
};

function cloneNamedExpression(expression: NamedExpression): NamedExpression {
  return { ...expression };
}

function cloneNamedExpressionScope(
  scope: NamedExpressionScope
): NamedExpressionScope {
  return { ...scope };
}

function cloneNamedExpressionBucket(
  bucket: NamedExpressionBucket
): NamedExpressionBucket {
  return { ...bucket };
}

function cloneNamedExpressionEntryState(
  state: NamedExpressionEntryState | undefined
): NamedExpressionEntryState | undefined {
  return state
    ? {
        scope: cloneNamedExpressionScope(state.scope),
        expressionName: state.expressionName,
        expression: cloneNamedExpression(state.expression),
        index: state.index,
      }
    : undefined;
}

function cloneNamedExpressionBucketState(
  state: NamedExpressionBucketState | undefined
): NamedExpressionBucketState | undefined {
  return state
    ? {
        bucket: cloneNamedExpressionBucket(state.bucket),
        index: state.index,
      }
    : undefined;
}

function cloneNamedExpressionMutation(
  change: NamedExpressionMutation
): NamedExpressionMutation {
  if (change.kind === "named-expression") {
    return {
      kind: change.kind,
      before: cloneNamedExpressionEntryState(change.before),
      after: cloneNamedExpressionEntryState(change.after),
    };
  }
  return {
    kind: change.kind,
    before: cloneNamedExpressionBucketState(change.before),
    after: cloneNamedExpressionBucketState(change.after),
  };
}

function namedExpressionScopeKey(scope: NamedExpressionScope): string {
  switch (scope.type) {
    case "global":
      return JSON.stringify([scope.type]);
    case "workbook":
      return JSON.stringify([scope.type, scope.workbookName]);
    case "sheet":
      return JSON.stringify([scope.type, scope.workbookName, scope.sheetName]);
  }
}

function namedExpressionEntryKey(state: NamedExpressionEntryState): string {
  return JSON.stringify([
    namedExpressionScopeKey(state.scope),
    state.expressionName,
  ]);
}

function namedExpressionBucketKey(bucket: NamedExpressionBucket): string {
  switch (bucket.type) {
    case "workbook":
    case "sheet-workbook":
      return JSON.stringify([bucket.type, bucket.workbookName]);
    case "sheet":
      return JSON.stringify([
        bucket.type,
        bucket.workbookName,
        bucket.sheetName,
      ]);
  }
}

function namedExpressionsEqual(
  left: NamedExpression,
  right: NamedExpression
): boolean {
  return left.name === right.name && left.expression === right.expression;
}

function namedExpressionEntryStatesEqual(
  left: NamedExpressionEntryState,
  right: NamedExpressionEntryState
): boolean {
  return (
    namedExpressionScopeKey(left.scope) ===
      namedExpressionScopeKey(right.scope) &&
    left.expressionName === right.expressionName &&
    left.index === right.index &&
    namedExpressionsEqual(left.expression, right.expression)
  );
}

function namedExpressionBucketStatesEqual(
  left: NamedExpressionBucketState,
  right: NamedExpressionBucketState
): boolean {
  return (
    namedExpressionBucketKey(left.bucket) ===
      namedExpressionBucketKey(right.bucket) && left.index === right.index
  );
}

type IndexedMapInsertion<TKey, TValue> = {
  key: TKey;
  value: TValue;
  index: number;
  order: number;
};

/** Rebuild an ordered Map once from absolute target indices. */
function applyOrderedMapPatch<TKey, TValue>(
  map: Map<TKey, TValue>,
  removedKeys: ReadonlySet<TKey>,
  rawInsertions: readonly IndexedMapInsertion<TKey, TValue>[]
): void {
  if (removedKeys.size === 0 && rawInsertions.length === 0) {
    return;
  }
  const insertionKeys = new Set(rawInsertions.map(({ key }) => key));
  const remaining = Array.from(map.entries()).filter(
    ([key]) => !removedKeys.has(key) && !insertionKeys.has(key)
  );
  const insertions = [...rawInsertions].sort(
    (left, right) => left.index - right.index || right.order - left.order
  );
  const rebuilt: Array<[TKey, TValue]> = [];
  let remainingIndex = 0;

  for (const insertion of insertions) {
    const targetIndex = Math.max(0, insertion.index);
    while (rebuilt.length < targetIndex && remainingIndex < remaining.length) {
      rebuilt.push(remaining[remainingIndex++]!);
    }
    rebuilt.push([insertion.key, insertion.value]);
  }
  while (remainingIndex < remaining.length) {
    rebuilt.push(remaining[remainingIndex++]!);
  }

  map.clear();
  for (const [key, value] of rebuilt) {
    map.set(key, value);
  }
}

export class NamedExpressionManager {
  sheetExpressions: Map<string, Map<string, Map<string, NamedExpression>>> =
    new Map();
  workbookExpressions: Map<string, Map<string, NamedExpression>> = new Map();
  globalExpressions: Map<string, NamedExpression> = new Map();

  private mutationBatchDepth = 0;
  private pendingMutationPatches: NamedExpressionMutation[][] = [];
  private mutationReportingSuppressionDepth = 0;

  constructor(
    private mutationObserver?: NamedExpressionMutationObserver,
    private readonly shouldObserve: () => boolean = () => true,
    private readonly detachMutationValues = true,
    private readonly shouldBatchMutations: () => boolean = () => true
  ) {}

  private get observingMutations(): boolean {
    return (
      this.mutationObserver !== undefined &&
      this.mutationReportingSuppressionDepth === 0 &&
      this.shouldObserve()
    );
  }

  batchMutations<T>(callback: () => T): T {
    if (!this.shouldBatchMutations()) {
      return callback();
    }

    this.mutationBatchDepth++;
    try {
      return callback();
    } finally {
      this.mutationBatchDepth--;
      if (this.mutationBatchDepth === 0) {
        const patches = this.pendingMutationPatches;
        this.pendingMutationPatches = [];
        for (const patch of patches) {
          this.emitMutationPatch(patch);
        }
      }
    }
  }

  private reportMutations(changes: readonly NamedExpressionMutation[]): void {
    if (!this.observingMutations || changes.length === 0) {
      return;
    }

    const detached = this.detachMutationValues
      ? changes.map(cloneNamedExpressionMutation)
      : [...changes];
    if (this.mutationBatchDepth > 0) {
      this.pendingMutationPatches.push(detached);
      return;
    }

    this.emitMutationPatch(detached);
  }

  private emitMutationPatch(changes: readonly NamedExpressionMutation[]): void {
    if (!this.observingMutations) {
      return;
    }
    this.mutationObserver!(changes);
  }

  private captureExpressionMap(
    scope: NamedExpressionScope,
    expressions: Map<string, NamedExpression>
  ): NamedExpressionEntryState[] {
    return Array.from(expressions, ([expressionName, expression], index) => ({
      scope: cloneNamedExpressionScope(scope),
      expressionName,
      expression: cloneNamedExpression(expression),
      index,
    }));
  }

  private captureState(): NamedExpressionState {
    const entries = this.captureExpressionMap(
      { type: "global" },
      this.globalExpressions
    );
    const buckets: NamedExpressionBucketState[] = [];

    let workbookIndex = 0;
    for (const [workbookName, expressions] of this.workbookExpressions) {
      buckets.push({
        bucket: { type: "workbook", workbookName },
        index: workbookIndex++,
      });
      entries.push(
        ...this.captureExpressionMap(
          { type: "workbook", workbookName },
          expressions
        )
      );
    }

    let sheetWorkbookIndex = 0;
    for (const [workbookName, sheets] of this.sheetExpressions) {
      buckets.push({
        bucket: { type: "sheet-workbook", workbookName },
        index: sheetWorkbookIndex++,
      });
      let sheetIndex = 0;
      for (const [sheetName, expressions] of sheets) {
        buckets.push({
          bucket: { type: "sheet", workbookName, sheetName },
          index: sheetIndex++,
        });
        entries.push(
          ...this.captureExpressionMap(
            { type: "sheet", workbookName, sheetName },
            expressions
          )
        );
      }
    }

    return { entries, buckets };
  }

  private captureWorkbookState(
    workbookNames: ReadonlySet<string>
  ): NamedExpressionState {
    const entries: NamedExpressionEntryState[] = [];
    const buckets: NamedExpressionBucketState[] = [];

    let workbookIndex = 0;
    for (const [workbookName, expressions] of this.workbookExpressions) {
      if (workbookNames.has(workbookName)) {
        buckets.push({
          bucket: { type: "workbook", workbookName },
          index: workbookIndex,
        });
        entries.push(
          ...this.captureExpressionMap(
            { type: "workbook", workbookName },
            expressions
          )
        );
      }
      workbookIndex++;
    }

    let sheetWorkbookIndex = 0;
    for (const [workbookName, sheets] of this.sheetExpressions) {
      if (workbookNames.has(workbookName)) {
        buckets.push({
          bucket: { type: "sheet-workbook", workbookName },
          index: sheetWorkbookIndex,
        });
        let sheetIndex = 0;
        for (const [sheetName, expressions] of sheets) {
          buckets.push({
            bucket: { type: "sheet", workbookName, sheetName },
            index: sheetIndex++,
          });
          entries.push(
            ...this.captureExpressionMap(
              { type: "sheet", workbookName, sheetName },
              expressions
            )
          );
        }
      }
      sheetWorkbookIndex++;
    }

    return { entries, buckets };
  }

  private captureSheetState(
    workbookName: string,
    sheetNames: ReadonlySet<string>
  ): NamedExpressionState {
    const entries: NamedExpressionEntryState[] = [];
    const buckets: NamedExpressionBucketState[] = [];
    const sheets = this.sheetExpressions.get(workbookName);
    if (!sheets) {
      return { entries, buckets };
    }

    let sheetIndex = 0;
    for (const [sheetName, expressions] of sheets) {
      if (sheetNames.has(sheetName)) {
        buckets.push({
          bucket: { type: "sheet", workbookName, sheetName },
          index: sheetIndex,
        });
        entries.push(
          ...this.captureExpressionMap(
            { type: "sheet", workbookName, sheetName },
            expressions
          )
        );
      }
      sheetIndex++;
    }
    return { entries, buckets };
  }

  private buildStateMutations(
    before: NamedExpressionState,
    after: NamedExpressionState
  ): NamedExpressionMutation[] {
    const mutations: NamedExpressionMutation[] = [];
    const beforeBuckets = new Map(
      before.buckets.map((state) => [
        namedExpressionBucketKey(state.bucket),
        state,
      ])
    );
    const afterBuckets = new Map(
      after.buckets.map((state) => [
        namedExpressionBucketKey(state.bucket),
        state,
      ])
    );
    const bucketKeys = new Set([
      ...beforeBuckets.keys(),
      ...afterBuckets.keys(),
    ]);
    for (const key of bucketKeys) {
      const beforeState = beforeBuckets.get(key);
      const afterState = afterBuckets.get(key);
      if (
        beforeState &&
        afterState &&
        namedExpressionBucketStatesEqual(beforeState, afterState)
      ) {
        continue;
      }
      mutations.push({
        kind: "named-expression-bucket",
        before: cloneNamedExpressionBucketState(beforeState),
        after: cloneNamedExpressionBucketState(afterState),
      });
    }

    const beforeEntries = new Map(
      before.entries.map((state) => [namedExpressionEntryKey(state), state])
    );
    const afterEntries = new Map(
      after.entries.map((state) => [namedExpressionEntryKey(state), state])
    );
    const entryKeys = new Set([
      ...beforeEntries.keys(),
      ...afterEntries.keys(),
    ]);
    for (const key of entryKeys) {
      const beforeState = beforeEntries.get(key);
      const afterState = afterEntries.get(key);
      if (
        beforeState &&
        afterState &&
        namedExpressionEntryStatesEqual(beforeState, afterState)
      ) {
        continue;
      }
      mutations.push({
        kind: "named-expression",
        before: cloneNamedExpressionEntryState(beforeState),
        after: cloneNamedExpressionEntryState(afterState),
      });
    }
    return mutations;
  }

  private mutateWithStateDiff<T>(callback: () => T): T {
    if (!this.observingMutations) {
      return callback();
    }
    const before = this.captureState();
    const result = callback();
    const after = this.captureState();
    this.reportMutations(this.buildStateMutations(before, after));
    return result;
  }

  private getExpressionMap(
    scope: NamedExpressionScope
  ): Map<string, NamedExpression> | undefined {
    switch (scope.type) {
      case "global":
        return this.globalExpressions;
      case "workbook":
        return this.workbookExpressions.get(scope.workbookName);
      case "sheet":
        return this.sheetExpressions
          .get(scope.workbookName)
          ?.get(scope.sheetName);
    }
  }

  private ensureExpressionMap(
    scope: NamedExpressionScope
  ): Map<string, NamedExpression> {
    if (scope.type === "global") {
      return this.globalExpressions;
    }
    if (scope.type === "workbook") {
      let expressions = this.workbookExpressions.get(scope.workbookName);
      if (!expressions) {
        expressions = new Map();
        this.workbookExpressions.set(scope.workbookName, expressions);
      }
      return expressions;
    }

    let sheets = this.sheetExpressions.get(scope.workbookName);
    if (!sheets) {
      sheets = new Map();
      this.sheetExpressions.set(scope.workbookName, sheets);
    }
    let expressions = sheets.get(scope.sheetName);
    if (!expressions) {
      expressions = new Map();
      sheets.set(scope.sheetName, expressions);
    }
    return expressions;
  }

  private getMapIndex<TKey, TValue>(map: Map<TKey, TValue>, key: TKey): number {
    let index = 0;
    for (const candidate of map.keys()) {
      if (Object.is(candidate, key)) {
        return index;
      }
      index++;
    }
    return -1;
  }

  private getScope(options: {
    sheetName?: string;
    workbookName?: string;
  }): NamedExpressionScope {
    if (options.sheetName && !options.workbookName) {
      throw new Error("Missing workbookName");
    }
    if (options.sheetName && options.workbookName) {
      return {
        type: "sheet",
        workbookName: options.workbookName,
        sheetName: options.sheetName,
      };
    }
    if (options.workbookName) {
      return { type: "workbook", workbookName: options.workbookName };
    }
    return { type: "global" };
  }

  private captureEntryState(
    scope: NamedExpressionScope,
    expressionName: string
  ): NamedExpressionEntryState | undefined {
    const map = this.getExpressionMap(scope);
    const expression = map?.get(expressionName);
    if (!map || !expression) {
      return undefined;
    }
    return {
      scope: cloneNamedExpressionScope(scope),
      expressionName,
      expression: cloneNamedExpression(expression),
      index: this.getMapIndex(map, expressionName),
    };
  }

  private captureBucketState(
    bucket: NamedExpressionBucket
  ): NamedExpressionBucketState | undefined {
    switch (bucket.type) {
      case "workbook": {
        if (!this.workbookExpressions.has(bucket.workbookName)) {
          return undefined;
        }
        return {
          bucket: cloneNamedExpressionBucket(bucket),
          index: this.getMapIndex(
            this.workbookExpressions,
            bucket.workbookName
          ),
        };
      }
      case "sheet-workbook": {
        if (!this.sheetExpressions.has(bucket.workbookName)) {
          return undefined;
        }
        return {
          bucket: cloneNamedExpressionBucket(bucket),
          index: this.getMapIndex(this.sheetExpressions, bucket.workbookName),
        };
      }
      case "sheet": {
        const sheets = this.sheetExpressions.get(bucket.workbookName);
        if (!sheets?.has(bucket.sheetName)) {
          return undefined;
        }
        return {
          bucket: cloneNamedExpressionBucket(bucket),
          index: this.getMapIndex(sheets, bucket.sheetName),
        };
      }
    }
  }

  private hasBucket(bucket: NamedExpressionBucket): boolean {
    switch (bucket.type) {
      case "workbook":
        return this.workbookExpressions.has(bucket.workbookName);
      case "sheet-workbook":
        return this.sheetExpressions.has(bucket.workbookName);
      case "sheet":
        return (
          this.sheetExpressions
            .get(bucket.workbookName)
            ?.has(bucket.sheetName) === true
        );
    }
  }

  private getScopeBuckets(
    scope: NamedExpressionScope
  ): NamedExpressionBucket[] {
    switch (scope.type) {
      case "global":
        return [];
      case "workbook":
        return [{ type: "workbook", workbookName: scope.workbookName }];
      case "sheet":
        return [
          { type: "sheet-workbook", workbookName: scope.workbookName },
          {
            type: "sheet",
            workbookName: scope.workbookName,
            sheetName: scope.sheetName,
          },
        ];
    }
  }

  private applyBucketMutations(
    changes: readonly Extract<
      NamedExpressionMutation,
      { kind: "named-expression-bucket" }
    >[],
    direction: "before" | "after"
  ): void {
    const workbookRemovals = new Set<string>();
    const sheetWorkbookRemovals = new Set<string>();
    const sheetRemovals = new Map<string, Set<string>>();
    const workbookInsertions: IndexedMapInsertion<
      string,
      Map<string, NamedExpression>
    >[] = [];
    const sheetWorkbookInsertions: IndexedMapInsertion<
      string,
      Map<string, Map<string, NamedExpression>>
    >[] = [];
    const sheetInsertions = new Map<
      string,
      IndexedMapInsertion<string, Map<string, NamedExpression>>[]
    >();

    const removeBucket = (bucket: NamedExpressionBucket): void => {
      switch (bucket.type) {
        case "workbook":
          workbookRemovals.add(bucket.workbookName);
          return;
        case "sheet-workbook":
          sheetWorkbookRemovals.add(bucket.workbookName);
          return;
        case "sheet": {
          let removals = sheetRemovals.get(bucket.workbookName);
          if (!removals) {
            removals = new Set();
            sheetRemovals.set(bucket.workbookName, removals);
          }
          removals.add(bucket.sheetName);
        }
      }
    };

    changes.forEach((change, order) => {
      if (change.before) {
        removeBucket(change.before.bucket);
      }
      if (change.after) {
        removeBucket(change.after.bucket);
      }

      const target = change[direction];
      if (!target) {
        return;
      }
      const other = direction === "before" ? change.after : change.before;
      const { bucket, index } = target;
      switch (bucket.type) {
        case "workbook": {
          const sourceName =
            other?.bucket.type === "workbook"
              ? other.bucket.workbookName
              : undefined;
          workbookInsertions.push({
            key: bucket.workbookName,
            value:
              this.workbookExpressions.get(bucket.workbookName) ??
              (sourceName === undefined
                ? undefined
                : this.workbookExpressions.get(sourceName)) ??
              new Map(),
            index,
            order,
          });
          return;
        }
        case "sheet-workbook": {
          const sourceName =
            other?.bucket.type === "sheet-workbook"
              ? other.bucket.workbookName
              : undefined;
          sheetWorkbookInsertions.push({
            key: bucket.workbookName,
            value:
              this.sheetExpressions.get(bucket.workbookName) ??
              (sourceName === undefined
                ? undefined
                : this.sheetExpressions.get(sourceName)) ??
              new Map(),
            index,
            order,
          });
          return;
        }
        case "sheet": {
          let insertions = sheetInsertions.get(bucket.workbookName);
          if (!insertions) {
            insertions = [];
            sheetInsertions.set(bucket.workbookName, insertions);
          }
          const currentSheets = this.sheetExpressions.get(bucket.workbookName);
          const sourceSheetName =
            other?.bucket.type === "sheet" &&
            other.bucket.workbookName === bucket.workbookName
              ? other.bucket.sheetName
              : undefined;
          insertions.push({
            key: bucket.sheetName,
            value:
              currentSheets?.get(bucket.sheetName) ??
              (sourceSheetName === undefined
                ? undefined
                : currentSheets?.get(sourceSheetName)) ??
              new Map(),
            index,
            order,
          });
        }
      }
    });

    applyOrderedMapPatch(
      this.workbookExpressions,
      workbookRemovals,
      workbookInsertions
    );
    applyOrderedMapPatch(
      this.sheetExpressions,
      sheetWorkbookRemovals,
      sheetWorkbookInsertions
    );

    for (const workbookName of new Set([
      ...sheetRemovals.keys(),
      ...sheetInsertions.keys(),
    ])) {
      let sheets = this.sheetExpressions.get(workbookName);
      const insertions = sheetInsertions.get(workbookName) ?? [];
      if (!sheets) {
        if (insertions.length === 0) {
          continue;
        }
        sheets = new Map();
        this.sheetExpressions.set(workbookName, sheets);
      }
      applyOrderedMapPatch(
        sheets,
        sheetRemovals.get(workbookName) ?? new Set(),
        insertions
      );
    }
  }

  /** Applies detached observer deltas exactly, including map order and buckets. */
  private applyMutations(
    changes: readonly NamedExpressionMutation[],
    direction: "before" | "after"
  ): void {
    this.batchMutations(() => {
      const entryChanges = changes.filter(
        (
          change
        ): change is Extract<
          NamedExpressionMutation,
          { kind: "named-expression" }
        > => change.kind === "named-expression"
      );
      const bucketChanges = changes.filter(
        (
          change
        ): change is Extract<
          NamedExpressionMutation,
          { kind: "named-expression-bucket" }
        > => change.kind === "named-expression-bucket"
      );

      for (const change of entryChanges) {
        for (const state of [change.before, change.after]) {
          if (state) {
            this.getExpressionMap(state.scope)?.delete(state.expressionName);
          }
        }
      }

      this.applyBucketMutations(bucketChanges, direction);

      const targetEntriesByScope = new Map<
        string,
        {
          scope: NamedExpressionScope;
          insertions: IndexedMapInsertion<string, NamedExpression>[];
        }
      >();
      entryChanges.forEach((change, order) => {
        const state = change[direction];
        if (!state) {
          return;
        }
        const scopeKey = namedExpressionScopeKey(state.scope);
        let target = targetEntriesByScope.get(scopeKey);
        if (!target) {
          target = { scope: state.scope, insertions: [] };
          targetEntriesByScope.set(scopeKey, target);
        }
        target.insertions.push({
          key: state.expressionName,
          value: cloneNamedExpression(state.expression),
          index: state.index,
          order,
        });
      });
      for (const { scope, insertions } of targetEntriesByScope.values()) {
        applyOrderedMapPatch(
          this.ensureExpressionMap(scope),
          new Set(),
          insertions
        );
      }

      const reported =
        direction === "after"
          ? changes
          : changes.map(
              (change): NamedExpressionMutation =>
                change.kind === "named-expression"
                  ? {
                      kind: change.kind,
                      before: change.after,
                      after: change.before,
                    }
                  : {
                      kind: change.kind,
                      before: change.after,
                      after: change.before,
                    }
            );
      this.reportMutations(reported);
    });
  }

  /** Applies retained deltas directly without notifying the observer. */
  applyHistoryChanges(
    changes: readonly NamedExpressionMutation[],
    direction: MutationDirection
  ): void {
    this.mutationReportingSuppressionDepth++;
    try {
      this.applyMutations(changes, direction === "undo" ? "before" : "after");
    } finally {
      this.mutationReportingSuppressionDepth--;
    }
  }

  addNamedExpression({
    expression,
    expressionName,
    sheetName,
    workbookName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    const scope = this.getScope({ sheetName, workbookName });
    const shouldObserve = this.observingMutations;
    const beforeEntry = shouldObserve
      ? this.captureEntryState(scope, expressionName)
      : undefined;
    const buckets = this.getScopeBuckets(scope);
    const missingBuckets = shouldObserve
      ? buckets.filter((bucket) => !this.hasBucket(bucket))
      : [];
    const namedExpression: NamedExpression = {
      name: expressionName,
      expression,
    };
    const expressions = this.ensureExpressionMap(scope);
    expressions.set(expressionName, namedExpression);

    const changes: NamedExpressionMutation[] = [];
    for (const bucket of missingBuckets) {
      const after = this.captureBucketState(bucket);
      changes.push({
        kind: "named-expression-bucket",
        before: undefined,
        after,
      });
    }
    const afterEntry = shouldObserve
      ? {
          scope: cloneNamedExpressionScope(scope),
          expressionName,
          expression: cloneNamedExpression(namedExpression),
          index: beforeEntry?.index ?? expressions.size - 1,
        }
      : undefined;
    if (
      !beforeEntry ||
      !afterEntry ||
      !namedExpressionEntryStatesEqual(beforeEntry, afterEntry)
    ) {
      changes.push({
        kind: "named-expression",
        before: beforeEntry,
        after: afterEntry,
      });
    }
    this.reportMutations(changes);
  }

  removeNamedExpression({
    expressionName,
    sheetName,
    workbookName,
  }: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): boolean {
    const scope = this.getScope({ sheetName, workbookName });
    const before = this.observingMutations
      ? this.captureEntryState(scope, expressionName)
      : undefined;
    const found = this.getExpressionMap(scope)?.delete(expressionName) ?? false;
    if (found && before) {
      this.reportMutations([
        { kind: "named-expression", before, after: undefined },
      ]);
    }
    return found;
  }

  updateNamedExpression({
    expression,
    expressionName,
    sheetName,
    workbookName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    // Check if the named expression exists
    let exists = false;

    if (sheetName && !workbookName) {
      throw new Error("Missing workbookName");
    }

    if (sheetName && workbookName) {
      const wbLevel = this.sheetExpressions.get(workbookName);
      if (wbLevel) {
        const sheetLevel = wbLevel.get(sheetName);
        if (sheetLevel) {
          exists = sheetLevel.has(expressionName);
        }
      }
    } else if (workbookName) {
      const workbookNamedExpressions =
        this.workbookExpressions.get(workbookName);
      if (workbookNamedExpressions) {
        exists = workbookNamedExpressions.has(expressionName);
      }
    } else {
      exists = this.globalExpressions.has(expressionName);
    }

    if (!exists) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Update is the same as add for existing expressions
    this.addNamedExpression({
      expression,
      expressionName,
      sheetName,
      workbookName,
    });
  }

  renameNamedExpression({
    expressionName,
    sheetName,
    workbookName,
    newName,
  }: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
    newName: string;
  }): boolean {
    // Check if the named expression exists
    let targetMap: Map<string, NamedExpression> | undefined;
    const scope = this.getScope({ sheetName, workbookName });

    let isGlobal = false;

    if (sheetName && workbookName) {
      const wbLevel = this.sheetExpressions.get(workbookName);
      if (wbLevel) {
        const sheetLevel = wbLevel.get(sheetName);
        if (sheetLevel) {
          targetMap = sheetLevel;
        }
      }
    } else if (workbookName) {
      targetMap = this.workbookExpressions.get(workbookName);
    } else {
      targetMap = this.globalExpressions;
      isGlobal = true;
    }

    if (!targetMap || !targetMap.has(expressionName)) {
      throw new Error(`Named expression '${expressionName}' does not exist`);
    }

    // Check if the new name already exists
    if (targetMap.has(newName)) {
      throw new Error(`Named expression '${newName}' already exists`);
    }

    // Get the expression to rename
    const namedExpression = targetMap.get(expressionName)!;
    const before = this.observingMutations
      ? this.captureEntryState(scope, expressionName)
      : undefined;

    // Update the name and re-add with new name
    const updatedExpression = { ...namedExpression, name: newName };
    targetMap.set(newName, updatedExpression);
    targetMap.delete(expressionName);

    if (before) {
      this.reportMutations([
        {
          kind: "named-expression",
          before,
          after: {
            scope: cloneNamedExpressionScope(scope),
            expressionName: newName,
            expression: cloneNamedExpression(updatedExpression),
            index: targetMap.size - 1,
          },
        },
      ]);
    }

    return true;
  }

  updateAllNamedExpressions(
    updateCallback: (
      formula: string,
      scope: { workbookName?: string; sheetName?: string }
    ) => string
  ): string[] {
    return this.batchMutations(() => {
      const changed = new Set<string>();
      const mutations: NamedExpressionMutation[] = [];

      const update = (
        map: Map<string, NamedExpression>,
        opts: { workbookName?: string; sheetName?: string }
      ) => {
        const scope = this.getScope(opts);
        let index = 0;
        map.forEach((namedExpr, name) => {
          const updatedExpression = updateCallback(namedExpr.expression, opts);

          if (updatedExpression !== namedExpr.expression) {
            const updated = {
              ...namedExpr,
              expression: updatedExpression,
            };
            map.set(name, updated);
            mutations.push({
              kind: "named-expression",
              before: {
                scope: cloneNamedExpressionScope(scope),
                expressionName: name,
                expression: cloneNamedExpression(namedExpr),
                index,
              },
              after: {
                scope: cloneNamedExpressionScope(scope),
                expressionName: name,
                expression: cloneNamedExpression(updated),
                index,
              },
            });
            changed.add(
              getNamedExpressionResourceKey({
                expressionName: name,
                workbookName: opts.workbookName,
                sheetName: opts.sheetName,
              })
            );
          }
          index++;
        });
      };

      update(this.globalExpressions, {});

      this.workbookExpressions.forEach((workbookLevel, workbookName) => {
        update(workbookLevel, { workbookName });
      });

      this.sheetExpressions.forEach((wbLevel, workbookName) => {
        wbLevel.forEach((sheetLevel, sheetName) => {
          update(sheetLevel, { workbookName, sheetName });
        });
      });

      this.reportMutations(mutations);
      return Array.from(changed);
    });
  }

  /**
   * Replace all named expressions
   */
  setNamedExpressions(
    opts: (
      | {
          type: "global";
        }
      | {
          type: "sheet";
          sheetName: string;
          workbookName: string;
        }
      | {
          type: "workbook";
          workbookName: string;
        }
    ) & {
      expressions: Map<string, NamedExpression>;
    }
  ) {
    let map: Map<string, NamedExpression> | undefined;

    if (opts.type === "sheet") {
      map = this.sheetExpressions.get(opts.workbookName)?.get(opts.sheetName);
    } else if (opts.type === "workbook") {
      map = this.workbookExpressions.get(opts.workbookName);
    } else {
      map = this.globalExpressions;
    }

    if (!map) {
      throw new Error("Invalid options: " + JSON.stringify(opts));
    }
    const scope: NamedExpressionScope =
      opts.type === "global"
        ? { type: "global" }
        : opts.type === "workbook"
        ? { type: "workbook", workbookName: opts.workbookName }
        : {
            type: "sheet",
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          };
    const before = this.observingMutations
      ? this.captureExpressionMap(scope, map)
      : [];

    this.batchMutations(() => {
      map.clear();
      opts.expressions.forEach((expression, name) => {
        map.set(name, expression);
      });
      if (this.observingMutations) {
        const after = this.captureExpressionMap(scope, map);
        this.reportMutations(
          this.buildStateMutations(
            { entries: before, buckets: [] },
            { entries: after, buckets: [] }
          )
        );
      }
    });
  }

  getNamedExpression(depNode: {
    name: string;
    scope:
      | {
          type: "global";
        }
      | {
          type: "workbook";
          workbookName: string;
        }
      | {
          type: "sheet";
          workbookName: string;
          sheetName: string;
        };
  }): NamedExpression | undefined {
    if (depNode.scope.type === "global") {
      return this.globalExpressions.get(depNode.name);
    }
    if (depNode.scope.type === "workbook") {
      return this.workbookExpressions
        .get(depNode.scope.workbookName)
        ?.get(depNode.name);
    }
    if (depNode.scope.type === "sheet") {
      return this.sheetExpressions
        .get(depNode.scope.workbookName)
        ?.get(depNode.scope.sheetName)
        ?.get(depNode.name);
    }
    return undefined;
  }

  resolveNamedExpressionWithScope(
    namedExpression: Pick<
      NamedExpressionNode,
      "name" | "sheetName" | "workbookName"
    >,
    context: EvaluationContext
  ):
    | {
        expression: string;
        scope:
          | { type: "global" }
          | { type: "workbook"; workbookName: string }
          | { type: "sheet"; workbookName: string; sheetName: string };
      }
    | undefined {
    // scenario 1: no sheetName nor workbookName
    if (!namedExpression.sheetName && !namedExpression.workbookName) {
      /**
       * the result of this formula will differ based on in which sheet and workbook we are
       * evaluating it in.
       */
      context.addContextDependency("sheet", "workbook");

      // step 1, check if there is a named expression in the sheet scope
      const expression = this.sheetExpressions
        .get(context.cellAddress.workbookName)
        ?.get(context.cellAddress.sheetName)
        ?.get(namedExpression.name);
      if (expression) {
        return {
          expression: expression.expression,
          scope: {
            type: "sheet",
            workbookName: context.cellAddress.workbookName,
            sheetName: context.cellAddress.sheetName,
          },
        };
      } else {
        // step 2, check if there is a named expression in the workbook scope
        const expression = this.workbookExpressions
          .get(context.cellAddress.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return {
            expression: expression.expression,
            scope: {
              type: "workbook",
              workbookName: context.cellAddress.workbookName,
            },
          };
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return {
              expression: expression.expression,
              scope: {
                type: "global",
              },
            };
          }
        }
      }
    }

    // scenario 2: we only have a workbookName - a bit weird, but could happen
    if (namedExpression.workbookName && !namedExpression.sheetName) {
      // special case: if workbook is the current workbook, we should just resolve the named expression according to scenario 1
      if (namedExpression.workbookName === context.cellAddress.workbookName) {
        return this.resolveNamedExpressionWithScope(
          {
            name: namedExpression.name,
          },
          context
        );
      }

      const expression = this.workbookExpressions
        .get(namedExpression.workbookName)
        ?.get(namedExpression.name);
      if (expression) {
        // step 1, check if there is a named expression in the workbook scope
        return {
          expression: expression.expression,
          scope: {
            type: "workbook",
            workbookName: namedExpression.workbookName,
          },
        };
      } else {
        // step 2, check if there is a named expression in the global scope
        const expression = this.globalExpressions.get(namedExpression.name);
        if (expression) {
          return {
            expression: expression.expression,
            scope: {
              type: "global",
            },
          };
        }
      }
    }

    // scenario 3: we only have a sheetName
    if (namedExpression.sheetName && !namedExpression.workbookName) {
      const expression = this.sheetExpressions
        .get(context.cellAddress.workbookName)
        ?.get(namedExpression.sheetName)
        ?.get(namedExpression.name);
      /**
       * the result of this formula will differ based on in which workbook we are
       * evaluating it in.
       */
      context.addContextDependency("workbook");
      if (expression) {
        // step 1, check if there is a named expression in the current workbook against the sheet name
        return {
          expression: expression.expression,
          scope: {
            type: "sheet",
            workbookName: context.cellAddress.workbookName,
            sheetName: namedExpression.sheetName,
          },
        };
      } else {
        // step 2, check if there is a named expression in the current workbook has a workbook scoped named expression
        const expression = this.workbookExpressions
          .get(context.cellAddress.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return {
            expression: expression.expression,
            scope: {
              type: "workbook",
              workbookName: context.cellAddress.workbookName,
            },
          };
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return {
              expression: expression.expression,
              scope: {
                type: "global",
              },
            };
          }
        }
      }
    }

    // scenario 4: we have both sheetName and workbookName
    if (namedExpression.sheetName && namedExpression.workbookName) {
      const expression = this.sheetExpressions
        .get(namedExpression.workbookName)
        ?.get(namedExpression.sheetName)
        ?.get(namedExpression.name);
      if (expression) {
        // step 1, check if there is a named expression the the sheet scope
        return {
          expression: expression.expression,
          scope: {
            type: "sheet",
            workbookName: namedExpression.workbookName,
            sheetName: namedExpression.sheetName,
          },
        };
      } else {
        // step 2, check if there is a named expression in the workbook scope
        const expression = this.workbookExpressions
          .get(namedExpression.workbookName)
          ?.get(namedExpression.name);
        if (expression) {
          return {
            expression: expression.expression,
            scope: {
              type: "workbook",
              workbookName: namedExpression.workbookName,
            },
          };
        } else {
          // step 3, check if there is a named expression in the global scope
          const expression = this.globalExpressions.get(namedExpression.name);
          if (expression) {
            return {
              expression: expression.expression,
              scope: {
                type: "global",
              },
            };
          }
        }
      }
    }
  }

  resolveNamedExpression(
    namedExpression: Pick<
      NamedExpressionNode,
      "name" | "sheetName" | "workbookName"
    >,
    context: EvaluationContext
  ): string | undefined {
    return this.resolveNamedExpressionWithScope(namedExpression, context)
      ?.expression;
  }

  getNamedExpressions() {
    return {
      sheetExpressions: this.sheetExpressions,
      workbookExpressions: this.workbookExpressions,
      globalExpressions: this.globalExpressions,
    };
  }

  resetNamedExpressions(
    namedExpressions: ReturnType<typeof this.getNamedExpressions>
  ) {
    return this.batchMutations(() =>
      this.mutateWithStateDiff(() => {
        this.clearInternal();
        namedExpressions.globalExpressions.forEach((expression, name) => {
          this.globalExpressions.set(name, expression);
        });

        namedExpressions.workbookExpressions.forEach(
          (workbookExpressions, workbookName) => {
            this.workbookExpressions.set(
              workbookName,
              new Map(workbookExpressions)
            );
          }
        );

        namedExpressions.sheetExpressions.forEach((sheets, workbookName) => {
          const restoredSheets = new Map<
            string,
            Map<string, NamedExpression>
          >();
          sheets.forEach((sheetExpressions, sheetName) => {
            restoredSheets.set(sheetName, new Map(sheetExpressions));
          });
          this.sheetExpressions.set(workbookName, restoredSheets);
        });
      })
    );
  }

  toSnapshot(): NamedExpressionManagerSnapshot {
    return this.getNamedExpressions();
  }

  restoreFromSnapshot(snapshot: NamedExpressionManagerSnapshot) {
    this.resetNamedExpressions(snapshot);
  }

  clear() {
    return this.batchMutations(() =>
      this.mutateWithStateDiff(() => this.clearInternal())
    );
  }

  private clearInternal(): void {
    this.sheetExpressions.clear();
    this.workbookExpressions.clear();
    this.globalExpressions.clear();
  }

  /**
   * When adding a sheet, we need to initialize the new maps
   */
  addSheet(opts: { workbookName: string; sheetName: string }) {
    const names = new Set([opts.sheetName]);
    const before = this.observingMutations
      ? this.captureSheetState(opts.workbookName, names)
      : undefined;
    const wbLevel = this.sheetExpressions.get(opts.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    const sheetLevel = wbLevel.get(opts.sheetName);
    if (sheetLevel) {
      throw new Error("Sheet already exists");
    }
    wbLevel.set(opts.sheetName, new Map());
    if (before) {
      const after = this.captureSheetState(opts.workbookName, names);
      this.reportMutations(this.buildStateMutations(before, after));
    }
  }

  /**
   * When adding a workbook, we need to initialize the new maps
   */
  addWorkbook(workbookName: string) {
    const names = new Set([workbookName]);
    const before = this.observingMutations
      ? this.captureWorkbookState(names)
      : undefined;
    this.batchMutations(() => {
      this.sheetExpressions.set(workbookName, new Map());
      this.workbookExpressions.set(workbookName, new Map());
      if (before) {
        const after = this.captureWorkbookState(names);
        this.reportMutations(this.buildStateMutations(before, after));
      }
    });
  }

  /**
   * When removing a workbook, we need to remove the workbook from the sheet level
   */
  removeWorkbook(workbookName: string) {
    const names = new Set([workbookName]);
    const before = this.observingMutations
      ? this.captureWorkbookState(names)
      : undefined;
    this.batchMutations(() => {
      this.sheetExpressions.delete(workbookName);
      this.workbookExpressions.delete(workbookName);
      if (before) {
        const after = this.captureWorkbookState(names);
        this.reportMutations(this.buildStateMutations(before, after));
      }
    });
  }

  /**
   * When removing a sheet, we need to remove the sheet from the workbook level
   */
  removeSheet(opts: { workbookName: string; sheetName: string }) {
    const names = new Set([opts.sheetName]);
    const before = this.observingMutations
      ? this.captureSheetState(opts.workbookName, names)
      : undefined;
    const wbLevel = this.sheetExpressions.get(opts.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    wbLevel.delete(opts.sheetName);
    if (before) {
      const after = this.captureSheetState(opts.workbookName, names);
      this.reportMutations(this.buildStateMutations(before, after));
    }
  }

  /**
   * Rename a sheet's named expressions, mainly used when renaming a sheet
   */
  renameSheet(options: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    const names = new Set([options.sheetName, options.newSheetName]);
    const before = this.observingMutations
      ? this.captureSheetState(options.workbookName, names)
      : undefined;
    const wbLevel = this.sheetExpressions.get(options.workbookName);
    if (!wbLevel) {
      throw new Error("Workbook not found");
    }
    const sheetLevel = wbLevel.get(options.sheetName);
    if (!sheetLevel) {
      throw new Error("Sheet not found");
    }
    wbLevel.set(options.newSheetName, sheetLevel);
    wbLevel.delete(options.sheetName);
    if (before) {
      const after = this.captureSheetState(options.workbookName, names);
      this.reportMutations(this.buildStateMutations(before, after));
    }
  }

  renameWorkbook(opts: { workbookName: string; newWorkbookName: string }) {
    const names = new Set([opts.workbookName, opts.newWorkbookName]);
    const before = this.observingMutations
      ? this.captureWorkbookState(names)
      : undefined;
    try {
      const wbLevel = this.sheetExpressions.get(opts.workbookName);
      if (!wbLevel) {
        throw new Error("Workbook not found");
      }
      this.sheetExpressions.set(opts.newWorkbookName, wbLevel);
      this.sheetExpressions.delete(opts.workbookName);

      const wbScopedExpressions = this.workbookExpressions.get(
        opts.workbookName
      );
      if (!wbScopedExpressions) {
        throw new Error("Workbook not found");
      }
      this.workbookExpressions.set(opts.newWorkbookName, wbScopedExpressions);
      this.workbookExpressions.delete(opts.workbookName);
    } finally {
      if (before) {
        const after = this.captureWorkbookState(names);
        this.reportMutations(this.buildStateMutations(before, after));
      }
    }
  }
}
