import { spawn, type ChildProcess } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { test as base, type APIRequestContext } from '@playwright/test';

type Fixtures = { server: { proc: ChildProcess; token: string; dataDir: string } };

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const REPO_ROOT = resolve(__dirname, '..', '..');

export const test = base.extend<Fixtures>({
  // Worker-scoped + auto so the bun server boots once per worker, persists
  // across the worker's tests, and the chat/reload specs (which only
  // destructure `{ page }`) still get the server running before navigation.
  server: [
    async ({}, use) => {
      const dataDir = mkdtempSync(join(tmpdir(), 'autoo-e2e-'));
      const token = 'e2e-token';
      const env = {
        ...process.env,
        AUTOOFFICE_TOKEN: token,
        AUTOOFFICE_DATA_DIR: dataDir,
        AUTOOFFICE_TEST_PROVIDER: 'fake',
        AUTOOFFICE_NO_TLS: '1',
        NODE_ENV: 'development',
      };
      const bunBin = process.env.AUTOOFFICE_BUN_BIN ?? 'bun';
      const proc = spawn(bunBin, ['apps/server/src/index.ts'], {
        cwd: REPO_ROOT,
        env,
        stdio: 'inherit',
      });

      // wait for /health to come up
      await waitForHealth('http://localhost:47318/health');

      await use({ proc, token, dataDir });

      proc.kill('SIGINT');
      // Give bun a beat to release the port before the next worker boots.
      await new Promise((r) => setTimeout(r, 500));
      rmSync(dataDir, { recursive: true, force: true });
    },
    { auto: true, scope: 'worker' },
  ],
});

export async function configureFakeProvider(
  request: APIRequestContext,
  token: string,
): Promise<void> {
  const headers = { Authorization: `Bearer ${token}` };
  const existing = await request.get('/api/providers', { headers });
  const providers = (await existing.json()) as Array<{ id: string }>;
  let providerId = providers[0]?.id;
  if (!providerId) {
    const created = await request.post('/api/providers', {
      headers,
      data: { kind: 'lmstudio', label: 'E2E local model' },
    });
    if (!created.ok()) throw new Error(`could not create E2E provider: ${created.status()}`);
    providerId = ((await created.json()) as { id: string }).id;
  }
  const settings = await request.put('/api/settings', {
    headers,
    data: { selectedProviderId: providerId, selectedModelId: 'fake-1' },
  });
  if (!settings.ok()) throw new Error(`could not select E2E provider: ${settings.status()}`);
}

async function waitForHealth(url: string): Promise<void> {
  for (let i = 0; i < 60; i++) {
    try {
      const res = await fetch(url);
      if (res.ok) return;
    } catch {
      /* not yet */
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  throw new Error('server did not come up');
}
