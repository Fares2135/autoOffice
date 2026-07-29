import type { HostKind } from '../host/context.ts';

export interface ExecutionResult {
  success: boolean;
  output?: unknown;
  error?: string;
  stack?: string;
  logs?: string[];
  debugInfo?: {
    code?: string;
    errorLocation?: string;
    statement?: string;
    surroundingStatements?: string[];
    fullStatements?: string[];
    message?: string;
  };
}

type PendingExecution = {
  resolve: (result: ExecutionResult) => void;
  timer: ReturnType<typeof setTimeout>;
};

const READY_TIMEOUT_MS = 10_000;

/**
 * Runs generated Office.js in a sandboxed, opaque-origin iframe.
 *
 * The iframe deliberately has `allow-scripts` but not `allow-same-origin`, so
 * generated code cannot read the task pane's DOM, localStorage, bearer token,
 * or other application state. Communication is restricted to an authenticated
 * MessageChannel-like exchange: responses are accepted only from the current
 * iframe window and for an outstanding random request id.
 */
export class Sandbox {
  private iframe: HTMLIFrameElement | null = null;
  private readyPromise: Promise<void> | null = null;
  private resolveReady: (() => void) | null = null;
  private readyTimer: ReturnType<typeof setTimeout> | null = null;
  private pending = new Map<string, PendingExecution>();
  private listening = false;

  constructor(private readonly host: HostKind) {}

  init(): void {
    if (!this.listening) {
      window.addEventListener('message', this.handleMessage);
      this.listening = true;
    }
    if (!this.iframe) this.createFrame();
  }

  destroy(): void {
    this.failPending('Execution sandbox was closed.');
    this.clearReadyTimer();
    this.resolveReady = null;
    this.readyPromise = null;
    this.iframe?.remove();
    this.iframe = null;
    if (this.listening) {
      window.removeEventListener('message', this.handleMessage);
      this.listening = false;
    }
  }

  async execute(code: string, timeout = 30_000): Promise<ExecutionResult> {
    if (!code.trim()) {
      return { success: false, error: 'No Office.js code was provided.', logs: [] };
    }
    this.init();

    try {
      await this.readyPromise;
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : String(error),
        logs: [],
      };
    }

    const frameWindow = this.iframe?.contentWindow;
    if (!frameWindow) {
      return { success: false, error: 'Execution sandbox is unavailable.', logs: [] };
    }

    const id = crypto.randomUUID();
    return new Promise<ExecutionResult>((resolve) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        resolve({
          success: false,
          error: `Execution timed out after ${timeout}ms and the sandbox was reset.`,
          logs: [],
        });
        this.resetFrame('Execution sandbox was reset after a timeout.');
      }, timeout);

      this.pending.set(id, { resolve, timer });
      frameWindow.postMessage({ type: 'execute', id, code, host: this.host }, '*');
    });
  }

  private createFrame(): void {
    const iframe = document.createElement('iframe');
    iframe.hidden = true;
    iframe.setAttribute('aria-hidden', 'true');
    iframe.setAttribute('sandbox', 'allow-scripts');
    iframe.setAttribute('referrerpolicy', 'no-referrer');
    iframe.src = new URL(`${import.meta.env.BASE_URL}iframe.html`, window.location.origin).toString();

    this.readyPromise = new Promise<void>((resolve, reject) => {
      this.resolveReady = resolve;
      this.readyTimer = setTimeout(() => {
        this.resolveReady = null;
        reject(new Error('Execution sandbox did not become ready.'));
      }, READY_TIMEOUT_MS);
    });

    this.iframe = iframe;
    document.body.appendChild(iframe);
  }

  private resetFrame(reason: string): void {
    this.failPending(reason);
    this.clearReadyTimer();
    this.resolveReady = null;
    this.readyPromise = null;
    this.iframe?.remove();
    this.iframe = null;
    this.createFrame();
  }

  private failPending(reason: string): void {
    for (const { resolve, timer } of this.pending.values()) {
      clearTimeout(timer);
      resolve({ success: false, error: reason, logs: [] });
    }
    this.pending.clear();
  }

  private clearReadyTimer(): void {
    if (this.readyTimer) clearTimeout(this.readyTimer);
    this.readyTimer = null;
  }

  private handleMessage = (event: MessageEvent): void => {
    if (!this.iframe || event.source !== this.iframe.contentWindow) return;
    const data = event.data as Partial<ExecutionResult> & { type?: string; id?: string };

    if (data?.type === 'sandbox-ready') {
      this.clearReadyTimer();
      this.resolveReady?.();
      this.resolveReady = null;
      return;
    }

    if ((data?.type !== 'result' && data?.type !== 'error') || !data.id) return;
    const pending = this.pending.get(data.id);
    if (!pending) return;

    clearTimeout(pending.timer);
    this.pending.delete(data.id);
    pending.resolve({
      success: data.type === 'result' && data.success === true,
      output: data.output,
      error: data.error,
      stack: data.stack,
      logs: Array.isArray(data.logs) ? data.logs : [],
      debugInfo: data.debugInfo,
    });
  };
}
