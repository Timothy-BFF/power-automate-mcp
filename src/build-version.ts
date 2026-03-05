// Cache-bust marker — forces Docker to recompile src/ layer
export const BUILD_VERSION = '1.0.3';
export const BUILD_TIMESTAMP = '2026-03-05T19:35:00Z';
// Fix: switch all Flow API calls to admin-scoped endpoints
// /scopes/admin/environments/{id}/v2/flows instead of /environments/{id}/v2/flows
