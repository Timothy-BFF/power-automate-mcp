// Cache-bust marker — forces Docker to recompile src/ layer
export const BUILD_VERSION = '1.0.5';
export const BUILD_TIMESTAMP = '2026-03-05T21:30:00Z';
// Fix: return raw arrays from listFlows/getFlowRuns (handlers call .length/.map)
// Fix: add getFlowRunDetails() alias (handlers call this, not getRunDetails)
// Fix: all return types Promise<any> to prevent type narrowing conflicts
