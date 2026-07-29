import { defineConfig } from 'vitest/config';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export default defineConfig({
  plugins: [
    {
      name: 'sql-as-text',
      enforce: 'pre',
      load(id) {
        if (!id.endsWith('.sql')) return null;
        return `export default ${JSON.stringify(readFileSync(id, 'utf8'))};`;
      },
    },
  ],
  test: {
    environment: 'node',
    include: ['src/**/*.test.ts'],
    // bun:sqlite + bun:ffi modules misbehave under vitest's default worker pool
    // when run via `bun --bun run vitest`. Forks single-fork keeps the bun runtime
    // happy and is fast enough for this codebase.
    pool: 'forks',
    poolOptions: {
      forks: {
        singleFork: true,
      },
    },
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov', 'html'],
      thresholds: {
        lines: 80,
        branches: 70,
        functions: 80,
        statements: 80,
      },
    },
  },
  resolve: {
    alias: {
      // fileURLToPath, not .pathname: on Windows .pathname gives "/D:/…", which
      // vite-node rejects with "File URL path must be an absolute path".
      '@shared': fileURLToPath(new URL('../../packages/shared/src', import.meta.url)),
    },
  },
});
