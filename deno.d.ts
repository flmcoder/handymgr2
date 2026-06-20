/**
 * Minimal Deno type definitions for TypeScript compatibility
 * This allows TypeScript to recognize Deno globals used in afproxy/ code
 */

declare global {
  namespace Deno {
    const env: {
      get(name: string): string | undefined;
      set?(name: string, value: string): void;
    };
  }
}

declare module 'npm:*' {
  const value: any;
  export = value;
}

export {};
