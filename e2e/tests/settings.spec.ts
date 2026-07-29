import { expect } from '@playwright/test';
import { test } from '../fixtures/boot-server';

test('settings page opens and accepts adding a provider', async ({ page, server }) => {
  await page.goto('/');
  // Open settings — adapt the selector to whatever the UI uses (gear icon, etc.).
  await page.getByRole('button', { name: 'Settings', exact: true }).click();
  await expect(page.getByText('Provider').first()).toBeVisible();

  // Add a CLI-bridge provider via API directly (UI selectors may vary).
  const resp = await page.request.post('/api/providers', {
    data: { kind: 'claude-code', label: 'Test CC' },
    headers: { Authorization: `Bearer ${server.token}` },
  });
  expect(resp.status()).toBe(201);
});
