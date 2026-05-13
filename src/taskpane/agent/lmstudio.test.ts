import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { discoverLmStudioModels, clearLmStudioCache, normalizeBaseUrl } from './lmstudio.ts';

const fetchMock = vi.fn();

beforeEach(() => {
  fetchMock.mockReset();
  vi.stubGlobal('fetch', fetchMock);
  clearLmStudioCache();
});

afterEach(() => {
  vi.unstubAllGlobals();
});

function jsonResponse(body: unknown, init: { status?: number } = {}): Response {
  return new Response(JSON.stringify(body), {
    status: init.status ?? 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('normalizeBaseUrl', () => {
  it('strips trailing slashes', () => {
    expect(normalizeBaseUrl('http://localhost:1234/v1/').baseUrl).toBe('http://localhost:1234/v1');
  });

  it('produces a rootBase by removing the /v1 suffix', () => {
    expect(normalizeBaseUrl('http://localhost:1234/v1').rootBase).toBe('http://localhost:1234');
  });

  it('leaves a non-/v1 base alone for rootBase', () => {
    expect(normalizeBaseUrl('http://localhost:1234').rootBase).toBe('http://localhost:1234');
  });

  it('handles double trailing slash', () => {
    expect(normalizeBaseUrl('http://localhost:1234/v1//').baseUrl).toBe('http://localhost:1234/v1');
  });
});

describe('discoverLmStudioModels — happy paths', () => {
  it('parses /api/v1/models metadata (LM Studio 0.4.x shape)', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({
      models: [
        {
          type: 'llm',
          key: 'qwen2.5-coder-7b-instruct',
          display_name: 'Qwen2.5 Coder 7B',
          architecture: 'qwen2',
          quantization: { name: 'Q4_K_M', bits_per_weight: 4 },
          max_context_length: 32768,
          capabilities: { vision: false, trained_for_tool_use: true },
          loaded_instances: [{ identifier: 'abc' }],
        },
        {
          type: 'llm',
          key: 'llama-3.2-3b-instruct',
          architecture: 'llama',
          loaded_instances: [],
        },
        {
          type: 'embedding',
          key: 'text-embed-nomic',
          loaded_instances: [],
        },
      ],
    }));

    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.status).toBe('connected');
    // Embedding entry should be filtered out.
    expect(result.models).toHaveLength(2);
    // Loaded should sort first.
    expect(result.models[0]).toMatchObject({
      id: 'qwen2.5-coder-7b-instruct',
      displayName: 'Qwen2.5 Coder 7B',
      architecture: 'qwen2',
      quantization: 'Q4_K_M',
      contextLength: 32768,
      capabilities: ['tool_use'],
      state: 'loaded',
    });
    expect(result.models[1]).toMatchObject({
      id: 'llama-3.2-3b-instruct',
      state: 'not-loaded',
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0][0]).toBe('http://localhost:1234/api/v1/models');
  });

  it('still parses older { data: [...] } native shape with id field', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({
      data: [
        { id: 'legacy-model', arch: 'legacy', quantization: 'Q4_0', state: 'loaded' },
      ],
    }));
    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.models[0]).toMatchObject({
      id: 'legacy-model',
      architecture: 'legacy',
      quantization: 'Q4_0',
      state: 'loaded',
    });
  });

  it('falls back to /v1/models when /api/v1/models returns 404', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({ error: 'not found' }, { status: 404 }))
      .mockResolvedValueOnce(jsonResponse({
        data: [
          { id: 'gpt-oss-20b' },
          { id: 'llama-3.1-8b' },
        ],
      }));

    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.status).toBe('connected');
    // Sorted alphabetically since neither has loaded state.
    expect(result.models.map(m => m.id).sort()).toEqual(['gpt-oss-20b', 'llama-3.1-8b']);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[1][0]).toBe('http://localhost:1234/v1/models');
  });

  it('falls back to /v1/models when /api/v1/models 500s', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({}, { status: 500 }))
      .mockResolvedValueOnce(jsonResponse({ data: [{ id: 'model-a' }] }));

    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.status).toBe('connected');
    expect(result.models).toEqual([{ id: 'model-a' }]);
  });

  it('caches { data: [...] } as the OpenAI-compat shape on fallback', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({}, { status: 404 }))
      .mockResolvedValueOnce(jsonResponse({
        data: [
          { id: 'gpt-oss-20b', object: 'model', owned_by: 'organization_owner' },
          { id: 'qwen3.5-4b', object: 'model', owned_by: 'organization_owner' },
        ],
      }));
    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.status).toBe('connected');
    expect(result.models.map(m => m.id)).toEqual(['gpt-oss-20b', 'qwen3.5-4b']);
  });
});

describe('discoverLmStudioModels — unreachable paths', () => {
  it('returns unreachable when both endpoints fail with TypeError (CORS/connection refused)', async () => {
    fetchMock.mockRejectedValue(new TypeError('Failed to fetch'));

    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.status).toBe('unreachable');
    expect(result.models).toEqual([]);
    expect(result.error).toMatch(/network error/i);
  });

  it('returns unreachable when both endpoints return errors', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({}, { status: 500 }))
      .mockResolvedValueOnce(jsonResponse({}, { status: 502 }));

    const result = await discoverLmStudioModels('http://localhost:1234/v1');
    expect(result.status).toBe('unreachable');
    expect(result.error).toMatch(/HTTP 502/);
  });
});

describe('discoverLmStudioModels — abort behaviour', () => {
  it('lets caller-initiated AbortError propagate', async () => {
    const ctrl = new AbortController();
    fetchMock.mockImplementation((_url, init: RequestInit) => {
      return new Promise((_resolve, reject) => {
        init.signal?.addEventListener('abort', () => {
          reject(new DOMException('Aborted', 'AbortError'));
        });
      });
    });

    const promise = discoverLmStudioModels('http://localhost:1234/v1', { signal: ctrl.signal });
    ctrl.abort();
    await expect(promise).rejects.toThrow();
  });
});

describe('discoverLmStudioModels — caching', () => {
  it('returns cached result within TTL without a second fetch', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ data: [{ id: 'cached-model' }] }));

    const first = await discoverLmStudioModels('http://localhost:1234/v1');
    const second = await discoverLmStudioModels('http://localhost:1234/v1');

    expect(first).toEqual(second);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('bypasses cache when force=true', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({ data: [{ id: 'first' }] }))
      .mockResolvedValueOnce(jsonResponse({ data: [{ id: 'second' }] }));

    await discoverLmStudioModels('http://localhost:1234/v1');
    const refreshed = await discoverLmStudioModels('http://localhost:1234/v1', { force: true });

    expect(refreshed.models[0].id).toBe('second');
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('caches unreachable results too (to avoid hammering when down)', async () => {
    fetchMock.mockRejectedValue(new TypeError('Failed to fetch'));

    const first = await discoverLmStudioModels('http://localhost:1234/v1');
    const second = await discoverLmStudioModels('http://localhost:1234/v1');

    expect(first.status).toBe('unreachable');
    expect(second.status).toBe('unreachable');
    // 2 fetches for the first call (native + fallback), 0 for the second.
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
});

describe('discoverLmStudioModels — URL normalization', () => {
  it('hits /api/v1/models, not /v1/api/v1/models, when baseUrl ends in /v1', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ data: [{ id: 'm1' }] }));
    await discoverLmStudioModels('http://localhost:1234/v1');
    expect(fetchMock.mock.calls[0][0]).toBe('http://localhost:1234/api/v1/models');
  });

  it('hits /api/v1/models when baseUrl has no /v1 suffix', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ data: [{ id: 'm1' }] }));
    await discoverLmStudioModels('http://localhost:1234');
    expect(fetchMock.mock.calls[0][0]).toBe('http://localhost:1234/api/v1/models');
  });

  it('hits /v1/models in fallback (preserving the /v1 suffix on baseUrl)', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({}, { status: 404 }))
      .mockResolvedValueOnce(jsonResponse({ data: [{ id: 'm1' }] }));

    await discoverLmStudioModels('http://localhost:1234/v1');
    expect(fetchMock.mock.calls[1][0]).toBe('http://localhost:1234/v1/models');
  });
});
