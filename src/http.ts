import {
  API_KEY_ENV_VAR,
  DEFAULT_BASE_URL,
  DEFAULT_TIMEOUT_MS,
  getGlobalConfig,
} from './config.js';
import { AuthenticationError, MitosisApiError, fromResponse } from './errors.js';
import { readEnvVar, toBuffer, warnAboutBrowserApiKeyExposure } from './runtime.js';
import type { ClientOptions, FetchLike } from './types.js';

export interface RequestParams {
  path: string;
  method: 'GET' | 'POST' | 'PUT' | 'DELETE';
  options?: ClientOptions;
  body?: FormData;
  query?: Record<string, string>;
  /** Set to false for the public, unauthenticated endpoints. */
  requiresApiKey?: boolean;
}

interface ResolvedOptions {
  apiKey: string | undefined;
  baseUrl: string;
  fetchImpl: FetchLike | undefined;
  timeoutMs: number;
}

function resolveOptions(options?: ClientOptions): ResolvedOptions {
  const global = getGlobalConfig();
  const baseUrl = options?.baseUrl ?? global.baseUrl ?? DEFAULT_BASE_URL;
  return {
    apiKey: options?.apiKey ?? global.apiKey ?? readEnvVar(API_KEY_ENV_VAR),
    baseUrl: baseUrl.replace(/\/+$/, ''),
    fetchImpl: options?.fetch ?? global.fetch ?? globalThis.fetch,
    timeoutMs: options?.timeoutMs ?? global.timeoutMs ?? DEFAULT_TIMEOUT_MS,
  };
}

async function readErrorBody(response: Response): Promise<unknown> {
  const raw = await response.text().catch(() => '');
  if (!raw) {
    return undefined;
  }
  try {
    return JSON.parse(raw) as unknown;
  } catch {
    return raw;
  }
}

async function send(
  fetchImpl: FetchLike | undefined,
  url: string,
  init: RequestInit,
  timeoutMs: number,
  label: string,
): Promise<Response> {
  if (!fetchImpl) {
    throw new MitosisApiError(
      'No global fetch available. Use Node.js 18 or later, or pass a custom { fetch } implementation.',
      0,
    );
  }

  let response: Response;
  try {
    response = await fetchImpl(url, { ...init, signal: AbortSignal.timeout(timeoutMs) });
  } catch (cause) {
    if (cause instanceof Error && (cause.name === 'TimeoutError' || cause.name === 'AbortError')) {
      throw new MitosisApiError(`Request to ${label} timed out after ${timeoutMs}ms`, 0);
    }
    throw cause;
  }

  if (!response.ok) {
    throw fromResponse(response.status, await readErrorBody(response));
  }
  return response;
}

export async function request(params: RequestParams): Promise<Response> {
  const { apiKey, baseUrl, fetchImpl, timeoutMs } = resolveOptions(params.options);
  const requiresApiKey = params.requiresApiKey !== false;

  if (requiresApiKey) {
    warnAboutBrowserApiKeyExposure();
  }
  if (requiresApiKey && !apiKey) {
    throw new AuthenticationError(
      `No API key configured. Pass { apiKey } to the constructor, call configure({ apiKey }), ` +
        `or set the ${API_KEY_ENV_VAR} environment variable. ` +
        `Get a key at ${DEFAULT_BASE_URL}/dashboard`,
    );
  }

  const url = new URL(baseUrl + params.path);
  for (const [key, value] of Object.entries(params.query ?? {})) {
    url.searchParams.set(key, value);
  }

  const headers: Record<string, string> = {};
  if (requiresApiKey && apiKey) {
    headers['x-api-key'] = apiKey;
  }

  return send(
    fetchImpl,
    url.toString(),
    { method: params.method, headers, body: params.body },
    timeoutMs,
    params.path,
  );
}

/**
 * Fetches an arbitrary resource, without API key and without the API base URL, so relative URLs
 * resolve against the page the bundle is served from.
 */
export async function fetchResource(url: string, options?: ClientOptions): Promise<Response> {
  const { fetchImpl, timeoutMs } = resolveOptions(options);
  return send(fetchImpl, url, { method: 'GET' }, timeoutMs, url);
}

export async function requestJson<T>(params: RequestParams): Promise<T> {
  const response = await request(params);
  return (await response.json()) as T;
}

export async function requestText(params: RequestParams): Promise<string> {
  const response = await request(params);
  return await response.text();
}

export async function requestBuffer(params: RequestParams): Promise<Buffer> {
  const response = await request(params);
  return toBuffer(await response.arrayBuffer());
}
