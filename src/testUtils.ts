import type { FetchLike } from './types.js';

export interface RecordedCall {
  url: string;
  method: string;
  headers: Record<string, string>;
  form?: FormData;
}

export interface FetchStub {
  fetch: FetchLike;
  calls: RecordedCall[];
  lastCall(): RecordedCall;
}

type Responder = (call: RecordedCall) => Response | Promise<Response>;

/** Builds an injectable fetch that records every call and replays scripted responses. */
export function createFetchStub(responder: Responder | Response): FetchStub {
  const calls: RecordedCall[] = [];
  const respond: Responder = typeof responder === 'function' ? responder : () => responder;

  const fetch: FetchLike = async (url, init) => {
    const call: RecordedCall = {
      url,
      method: init?.method ?? 'GET',
      headers: (init?.headers as Record<string, string> | undefined) ?? {},
      form: init?.body instanceof FormData ? init.body : undefined,
    };
    calls.push(call);
    return respond(call);
  };

  return {
    fetch,
    calls,
    lastCall: () => calls[calls.length - 1],
  };
}

export function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

export function textResponse(body: string, status = 200): Response {
  return new Response(body, { status, headers: { 'content-type': 'text/plain' } });
}

export function binaryResponse(bytes: Uint8Array, status = 200): Response {
  return new Response(bytes, { status });
}

export async function formFileBytes(form: FormData, field = 'file'): Promise<Uint8Array> {
  const entry = form.get(field);
  if (!(entry instanceof Blob)) {
    throw new Error(`Form field "${field}" is not a file`);
  }
  return new Uint8Array(await entry.arrayBuffer());
}

export interface BrowserStub {
  /** Every download the client triggered, in order. */
  downloads: Array<{ fileName: string; url: string }>;
  restore(): void;
}

/** Installs the minimal `window` the client probes for, to exercise the browser branches on Node. */
export function fakeBrowser(): BrowserStub {
  const downloads: Array<{ fileName: string; url: string }> = [];
  const globals = globalThis as { window?: unknown };

  globals.window = {
    // A non local hostname keeps the development-only api key exposure warning quiet.
    location: { hostname: 'app.example.com' },
    document: {
      createElement: () => {
        const link = {
          href: '',
          download: '',
          click: () => downloads.push({ fileName: link.download, url: link.href }),
        };
        return link;
      },
    },
  };

  return {
    downloads,
    restore: () => {
      delete globals.window;
    },
  };
}
