import { afterEach, describe, expect, it, vi } from 'vitest';
import { Sandbox } from './sandbox.ts';

function currentFrame(): HTMLIFrameElement {
  const iframe = document.querySelector('iframe[title], iframe[aria-hidden="true"]');
  if (!(iframe instanceof HTMLIFrameElement)) throw new Error('sandbox iframe not found');
  return iframe;
}

function dispatchFrom(frame: HTMLIFrameElement, data: unknown): void {
  window.dispatchEvent(new MessageEvent('message', {
    source: frame.contentWindow,
    data,
  }));
}

describe('Sandbox iframe bridge', () => {
  const sandboxes: Sandbox[] = [];

  afterEach(() => {
    for (const sandbox of sandboxes) sandbox.destroy();
    sandboxes.length = 0;
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it('creates an opaque-origin, script-only iframe', () => {
    const sandbox = new Sandbox('word');
    sandboxes.push(sandbox);
    sandbox.init();

    const frame = currentFrame();
    expect(frame.getAttribute('sandbox')).toBe('allow-scripts');
    expect(frame.getAttribute('sandbox')).not.toContain('allow-same-origin');
    expect(frame.hidden).toBe(true);
    expect(new URL(frame.src).pathname).toMatch(/\/iframe\.html$/);
  });

  it('sends code only after readiness and resolves a matching response', async () => {
    const sandbox = new Sandbox('word');
    sandboxes.push(sandbox);
    sandbox.init();
    const frame = currentFrame();
    const postMessage = vi.spyOn(frame.contentWindow!, 'postMessage');

    dispatchFrom(frame, { type: 'sandbox-ready' });
    const resultPromise = sandbox.execute('return 42;');
    await vi.waitFor(() => expect(postMessage).toHaveBeenCalledOnce());

    const request = postMessage.mock.calls[0]![0] as {
      type: string; id: string; code: string; host: string;
    };
    expect(request).toMatchObject({ type: 'execute', code: 'return 42;', host: 'word' });

    dispatchFrom(frame, {
      type: 'result',
      id: request.id,
      success: true,
      output: 42,
      logs: ['ok'],
    });
    await expect(resultPromise).resolves.toEqual({
      success: true,
      output: 42,
      error: undefined,
      stack: undefined,
      logs: ['ok'],
      debugInfo: undefined,
    });
  });

  it('ignores messages from other windows and unknown request ids', async () => {
    const sandbox = new Sandbox('excel');
    sandboxes.push(sandbox);
    sandbox.init();
    const frame = currentFrame();
    const postMessage = vi.spyOn(frame.contentWindow!, 'postMessage');
    dispatchFrom(frame, { type: 'sandbox-ready' });

    const resultPromise = sandbox.execute('return 7;');
    await vi.waitFor(() => expect(postMessage).toHaveBeenCalledOnce());
    const request = postMessage.mock.calls[0]![0] as { id: string };

    window.dispatchEvent(new MessageEvent('message', {
      source: window,
      data: { type: 'result', id: request.id, success: true, output: 'spoofed' },
    }));
    dispatchFrom(frame, { type: 'result', id: 'unknown', success: true, output: 'spoofed' });
    dispatchFrom(frame, { type: 'result', id: request.id, success: true, output: 7 });

    await expect(resultPromise).resolves.toMatchObject({ success: true, output: 7 });
  });

  it('preserves Office debug information returned by the iframe', async () => {
    const sandbox = new Sandbox('word');
    sandboxes.push(sandbox);
    sandbox.init();
    const frame = currentFrame();
    const postMessage = vi.spyOn(frame.contentWindow!, 'postMessage');
    dispatchFrom(frame, { type: 'sandbox-ready' });

    const resultPromise = sandbox.execute('return context.document.body.text;');
    await vi.waitFor(() => expect(postMessage).toHaveBeenCalledOnce());
    const request = postMessage.mock.calls[0]![0] as { id: string };

    dispatchFrom(frame, {
      type: 'error',
      id: request.id,
      success: false,
      error: 'PropertyNotLoaded',
      debugInfo: { errorLocation: 'Body.text', statement: 'body.text' },
    });

    await expect(resultPromise).resolves.toMatchObject({
      success: false,
      error: 'PropertyNotLoaded',
      debugInfo: { errorLocation: 'Body.text', statement: 'body.text' },
    });
  });

  it('resets the iframe when execution times out', async () => {
    const sandbox = new Sandbox('powerpoint');
    sandboxes.push(sandbox);
    sandbox.init();
    const original = currentFrame();
    dispatchFrom(original, { type: 'sandbox-ready' });

    const result = await sandbox.execute('return new Promise(() => {});', 5);
    expect(result.success).toBe(false);
    expect(result.error).toContain('sandbox was reset');
    expect(currentFrame()).not.toBe(original);
  });
});
