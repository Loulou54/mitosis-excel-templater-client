import type { ClientOptions } from './types.js';

/** Default host of the hosted Mitosis Excel Templater API. */
export const DEFAULT_BASE_URL = 'https://www.mitosis-excel-templater.dev';

/** Environment variable read as a last-resort source for the API key. */
export const API_KEY_ENV_VAR = 'MITOSIS_EXCEL_TEMPLATER_API_KEY';

export const DEFAULT_TIMEOUT_MS = 60_000;

let globalConfig: ClientOptions = {};

/**
 * Sets application-wide defaults for every client in the process.
 * Options passed to a constructor still take precedence.
 */
export function configure(options: ClientOptions): void {
  globalConfig = { ...globalConfig, ...options };
}

export function getGlobalConfig(): ClientOptions {
  return globalConfig;
}

/** Clears the application-wide defaults. */
export function resetConfig(): void {
  globalConfig = {};
}
