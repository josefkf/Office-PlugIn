// Test-only replacement for src/auth.js (wired in by `node build.mjs --mock-auth`).
export function isAuthSupported() { return true; }
export async function getToken() { return "test-token"; }
