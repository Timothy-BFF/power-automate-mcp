// Cache-bust marker — forces Docker to recompile src/ layer
export const BUILD_VERSION = '1.0.4';
export const BUILD_TIMESTAMP = '2026-03-05T19:42:00Z';
// Fix: backward-compatible FlowClient signatures with any params
// Adds getFlowRuns method name, typed return values for noImplicitAny
// All endpoints use admin-scoped /scopes/admin/ paths
