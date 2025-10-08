type DebugFlags = {
  isProfiling: boolean;
  profilingNamespaces: Record<string, boolean>;
  numEvaluationCalls: number;
  profiledCall: number;
  maxEvaluationCalls: number;
};

const PROFILED_CALL = 943;
const MAX_EVALUATION_CALLS = PROFILED_CALL + 1;

export const flags: DebugFlags = {
  isProfiling: false,
  profilingNamespaces: {},
  numEvaluationCalls: 0,
  profiledCall: PROFILED_CALL,
  maxEvaluationCalls: MAX_EVALUATION_CALLS,
};
