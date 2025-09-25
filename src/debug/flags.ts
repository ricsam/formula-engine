type DebugFlags = {
  isProfiling: boolean;
  profilingNamespaces: Record<string, boolean>;
  numEvaluationCalls: number;
  profiledCall: number;
  maxEvaluationCalls: number;
};

export const flags: DebugFlags = {
  isProfiling: false,
  profilingNamespaces: {},
  numEvaluationCalls: 0,
  profiledCall: 282,
  maxEvaluationCalls: 282,
};
