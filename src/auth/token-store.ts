export interface StoredToken {
  accessToken: string;
  expiresAt: number;
}

export class TokenStore {
  private tokens: Map<string, StoredToken> = new Map();

  get(scope: string): string | null {
    const stored = this.tokens.get(scope);
    if (!stored) return null;
    if (Date.now() >= stored.expiresAt - 5 * 60 * 1000) return null;
    return stored.accessToken;
  }

  set(scope: string, accessToken: string, expiresInSeconds: number): void {
    this.tokens.set(scope, {
      accessToken,
      expiresAt: Date.now() + expiresInSeconds * 1000,
    });
  }

  clear(): void {
    this.tokens.clear();
  }
}
