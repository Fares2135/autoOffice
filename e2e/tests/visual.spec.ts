import { expect } from '@playwright/test';
import { test } from '../fixtures/boot-server';

test('Arabic dark onboarding is RTL, responsive, and free of console errors', async ({ page, server }) => {
  const headers = { Authorization: `Bearer ${server.token}` };
  const providersResponse = await page.request.get('/api/providers', { headers });
  const providers = (await providersResponse.json()) as Array<{ id: string }>;
  for (const provider of providers) {
    await page.request.delete(`/api/providers/${provider.id}`, { headers });
  }
  await page.request.put('/api/settings', {
    headers,
    data: {
      locale: 'ar',
      theme: 'dark',
      selectedProviderId: null,
      selectedModelId: null,
    },
  });

  const errors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  await page.setViewportSize({ width: 375, height: 760 });
  await page.goto('/');

  await expect(page.locator('html')).toHaveAttribute('dir', 'rtl');
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await expect(page.getByRole('heading', { name: /اربط مساحة عمل/ })).toBeVisible();
  const overflow = await page.evaluate(() => {
    const viewportWidth = document.documentElement.clientWidth;
    return [...document.querySelectorAll<HTMLElement>('body *')]
      .map((element) => {
        const rect = element.getBoundingClientRect();
        return {
          tag: element.tagName.toLowerCase(),
          label: element.getAttribute('aria-label') ?? element.textContent?.trim().slice(0, 60),
          left: Math.round(rect.left),
          right: Math.round(rect.right),
          width: Math.round(rect.width),
        };
      })
      .filter((item) => item.left < -1 || item.right > viewportWidth + 1);
  });
  expect(overflow, JSON.stringify(overflow, null, 2)).toEqual([]);
  expect(errors).toEqual([]);

  await page.screenshot({
    path: '../.agent/qa/arabic-dark-onboarding.png',
    fullPage: true,
  });
});
