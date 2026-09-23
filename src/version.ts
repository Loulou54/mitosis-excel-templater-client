import { requestJson } from './http.js';
import type { ClientOptions, EngineVersion } from './types.js';

/** Returns the version of the Excel templating engine powering the API. Requires no API key. */
export async function getEngineVersion(options?: ClientOptions): Promise<EngineVersion> {
  return requestJson<EngineVersion>({
    path: '/api/mitosis-excel-templater-version',
    method: 'GET',
    options,
    requiresApiKey: false,
  });
}
