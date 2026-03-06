import { Request, Response, NextFunction } from 'express';

/**
 * Simtheory Auth Middleware (Permissive Mode)
 *
 * SIMTHEORY_AUTH_TOKEN is validated during MCP registration by Simtheory,
 * not for per-request SSE/REST authentication. This middleware logs
 * auth attempts but never blocks requests.
 */
export function simtheoryAuth(req: Request, _res: Response, next: NextFunction): void {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.replace('Bearer ', '').trim();
    const expected = (process.env.SIMTHEORY_AUTH_TOKEN || '').trim();
    if (expected && token !== expected) {
      console.warn('[Auth] Simtheory token mismatch (permissive mode - allowing through)');
    }
  }
  next();
}
