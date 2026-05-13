/**
 * LM Studio model discovery.
 *
 * LM Studio exposes two list endpoints:
 *   1. `GET <root>/api/v1/models` — native API with rich metadata (architecture,
 *      quantization, context length, capabilities, loaded/not-loaded state).
 *   2. `GET <baseUrl>/v1/models`  — OpenAI-compatible shape, ids only.
 *
 * `baseUrl` is the OpenAI-compatible base (typically `http://localhost:1234/v1`),
 * which is what `createOpenAICompatible` consumes. For the native endpoint we
 * strip a trailing `/v1` so we hit `<host>/api/v1/models`, not `<host>/v1/api/v1/models`.
 */

export interface LmStudioModel {
  id: string;
  displayName?: string;
  architecture?: string;
  quantization?: string;
  contextLength?: number;
  capabilities?: string[];
  state?: 'loaded' | 'not-loaded';
}

export type LmStudioStatus = 'idle' | 'connecting' | 'connected' | 'unreachable';

export interface LmStudioDiscovery {
  status: LmStudioStatus;
  models: LmStudioModel[];
  error?: string;
}

interface DiscoveryOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
  force?: boolean;
}

const DEFAULT_TIMEOUT_MS = 2500;
const CACHE_TTL_MS = 30_000;

const cache = new Map<string, { result: LmStudioDiscovery; expiresAt: number }>();

export function normalizeBaseUrl(raw: string): { baseUrl: string; rootBase: string } {
  const trimmed = (raw || '').replace(/\/+$/, '');
  const rootBase = trimmed.endsWith('/v1') ? trimmed.slice(0, -3) : trimmed;
  return { baseUrl: trimmed, rootBase };
}

export function clearLmStudioCache(): void {
  cache.clear();
}

export async function discoverLmStudioModels(
  rawBaseUrl: string,
  opts: DiscoveryOptions = {},
): Promise<LmStudioDiscovery> {
  const { baseUrl, rootBase } = normalizeBaseUrl(rawBaseUrl);
  const cacheKey = baseUrl;
  const now = Date.now();

  if (!opts.force) {
    const hit = cache.get(cacheKey);
    if (hit && hit.expiresAt > now) return hit.result;
  }

  const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;

  // Try /api/v1/models first (rich metadata).
  const native = await tryFetch(`${rootBase}/api/v1/models`, opts.signal, timeoutMs);
  if (native.ok && native.body) {
    const models = parseNative(native.body);
    const result: LmStudioDiscovery = {
      status: 'connected',
      models: sortModels(models),
    };
    cache.set(cacheKey, { result, expiresAt: now + CACHE_TTL_MS });
    return result;
  }

  // Fall back to /v1/models (OpenAI shape).
  const compat = await tryFetch(`${baseUrl}/models`, opts.signal, timeoutMs);
  if (compat.ok && compat.body) {
    const models = parseOpenAi(compat.body);
    const result: LmStudioDiscovery = {
      status: 'connected',
      models: sortModels(models),
    };
    cache.set(cacheKey, { result, expiresAt: now + CACHE_TTL_MS });
    return result;
  }

  const result: LmStudioDiscovery = {
    status: 'unreachable',
    models: [],
    error: compat.error ?? native.error ?? 'Unknown error',
  };
  cache.set(cacheKey, { result, expiresAt: now + CACHE_TTL_MS });
  return result;
}

interface FetchAttempt {
  ok: boolean;
  body?: unknown;
  error?: string;
}

async function tryFetch(url: string, externalSignal: AbortSignal | undefined, timeoutMs: number): Promise<FetchAttempt> {
  const ctrl = new AbortController();
  const onAbort = () => ctrl.abort(externalSignal?.reason);
  if (externalSignal) {
    if (externalSignal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }
    externalSignal.addEventListener('abort', onAbort, { once: true });
  }
  const timer = setTimeout(() => ctrl.abort(new DOMException('Timeout', 'TimeoutError')), timeoutMs);

  try {
    const response = await fetch(url, {
      method: 'GET',
      headers: { Accept: 'application/json' },
      signal: ctrl.signal,
    });
    if (!response.ok) {
      return { ok: false, error: `HTTP ${response.status} at ${url}` };
    }
    const body = await response.json();
    return { ok: true, body };
  } catch (err) {
    // Let user-initiated aborts bubble up so callers can distinguish them.
    if (externalSignal?.aborted) {
      throw err;
    }
    if (err instanceof DOMException && err.name === 'AbortError') {
      // Internal timeout — treat as unreachable.
      return { ok: false, error: `Timeout after ${timeoutMs}ms at ${url}` };
    }
    if (err instanceof TypeError) {
      // CORS / DNS / connection refused all surface as `TypeError: Failed to fetch`.
      return { ok: false, error: `Network error at ${url}: ${err.message}` };
    }
    return { ok: false, error: `Fetch failed at ${url}: ${(err as Error).message}` };
  } finally {
    clearTimeout(timer);
    if (externalSignal) externalSignal.removeEventListener('abort', onAbort);
  }
}

function parseNative(body: unknown): LmStudioModel[] {
  // LM Studio 0.4.x uses `{ models: [...] }`; older / alternate builds use `{ data: [...] }`.
  const list = extractList(body, ['models', 'data']);
  return list.flatMap((raw): LmStudioModel[] => {
    const r = raw as Record<string, unknown>;
    // Skip non-LLM entries (embeddings, etc.) — they aren't usable as chat models.
    const modelType = stringField(r.type);
    if (modelType && modelType !== 'llm' && modelType !== 'vlm') return [];

    const id = stringField(r.key) ?? stringField(r.id) ?? stringField(r.model_key) ?? '';
    if (!id) return [];

    const loadedInstances = Array.isArray(r.loaded_instances) ? r.loaded_instances : undefined;
    const state: LmStudioModel['state'] =
      loadedInstances !== undefined ? (loadedInstances.length > 0 ? 'loaded' : 'not-loaded') :
      r.state === 'loaded' || r.loaded === true ? 'loaded' :
      r.state === 'not-loaded' || r.loaded === false ? 'not-loaded' :
      undefined;

    return [{
      id,
      displayName: stringField(r.display_name) ?? stringField(r.displayName),
      architecture: stringField(r.arch) ?? stringField(r.architecture),
      quantization: parseQuantization(r.quantization),
      contextLength: numberField(r.context_length) ?? numberField(r.max_context_length),
      capabilities: parseCapabilities(r.capabilities),
      state,
    }];
  });
}

function parseOpenAi(body: unknown): LmStudioModel[] {
  const list = extractList(body, ['data', 'models']);
  return list.map((raw): LmStudioModel => {
    const r = raw as Record<string, unknown>;
    const id = stringField(r.id) ?? stringField(r.key) ?? '';
    return { id };
  }).filter(m => m.id);
}

function extractList(body: unknown, keys: string[]): unknown[] {
  if (!body || typeof body !== 'object') return [];
  for (const key of keys) {
    const v = (body as Record<string, unknown>)[key];
    if (Array.isArray(v)) return v;
  }
  return [];
}

function parseQuantization(v: unknown): string | undefined {
  // Native API: object `{ name, bits_per_weight }`. Older builds: bare string.
  if (typeof v === 'string') return v.length > 0 ? v : undefined;
  if (v && typeof v === 'object') {
    const name = (v as Record<string, unknown>).name;
    if (typeof name === 'string' && name.length > 0) return name;
  }
  return undefined;
}

function parseCapabilities(v: unknown): string[] | undefined {
  // Native API: object `{ vision: bool, trained_for_tool_use: bool, reasoning?: {...} }`.
  // Older builds: string array.
  if (Array.isArray(v)) {
    const out = v.filter((x): x is string => typeof x === 'string');
    return out.length > 0 ? out : undefined;
  }
  if (v && typeof v === 'object') {
    const obj = v as Record<string, unknown>;
    const caps: string[] = [];
    if (obj.vision === true) caps.push('vision');
    if (obj.trained_for_tool_use === true) caps.push('tool_use');
    if (obj.reasoning && typeof obj.reasoning === 'object') caps.push('reasoning');
    return caps.length > 0 ? caps : undefined;
  }
  return undefined;
}

function stringField(v: unknown): string | undefined {
  return typeof v === 'string' && v.length > 0 ? v : undefined;
}

function numberField(v: unknown): number | undefined {
  return typeof v === 'number' && Number.isFinite(v) ? v : undefined;
}

function sortModels(models: LmStudioModel[]): LmStudioModel[] {
  return [...models].sort((a, b) => {
    const aLoaded = a.state === 'loaded' ? 0 : 1;
    const bLoaded = b.state === 'loaded' ? 0 : 1;
    if (aLoaded !== bLoaded) return aLoaded - bLoaded;
    return a.id.localeCompare(b.id);
  });
}
