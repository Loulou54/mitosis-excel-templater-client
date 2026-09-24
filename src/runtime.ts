import { PRICING_URL } from './config.js';

/** Hostnames a browser bundle is served from while developing. */
const LOCAL_HOSTNAMES = new Set(['localhost', '127.0.0.1', '[::1]', '0.0.0.0', '']);

/** Browser globals, typed locally so the package compiles without the DOM lib. */
interface BrowserWindow {
  document?: BrowserDocument;
  location?: { hostname?: string };
}

interface BrowserDocument {
  createElement(tagName: string): BrowserAnchor;
  body?: { appendChild(node: unknown): void; removeChild(node: unknown): void };
}

interface BrowserAnchor {
  href: string;
  download: string;
  click(): void;
}

interface ObjectUrlFactory {
  createObjectURL(blob: Blob): string;
  revokeObjectURL(url: string): void;
}

function getWindow(): BrowserWindow | undefined {
  return (globalThis as { window?: BrowserWindow }).window;
}

function getProcess(): NodeJS.Process | undefined {
  return (globalThis as { process?: NodeJS.Process }).process;
}

/** Reads an environment variable without assuming `process` exists, so browsers do not blow up. */
export function readEnvVar(name: string): string | undefined {
  return getProcess()?.env?.[name];
}

export function isBrowser(): boolean {
  return getWindow()?.document !== undefined;
}

/**
 * Best effort development detection: `NODE_ENV` when a bundler or Node provides it, and the
 * hostname otherwise, so production bundles stay silent.
 */
export function isDevelopment(): boolean {
  const nodeEnv = readEnvVar('NODE_ENV');
  if (nodeEnv) {
    return nodeEnv !== 'production';
  }
  return LOCAL_HOSTNAMES.has(getWindow()?.location?.hostname ?? '');
}

/**
 * Converts raw response bytes to a Buffer on Node.js, and to a plain Uint8Array in the browser,
 * where `Buffer` does not exist.
 */
export function toBuffer(data: ArrayBuffer): Buffer {
  const bufferCtor = (globalThis as { Buffer?: typeof Buffer }).Buffer;
  return bufferCtor ? bufferCtor.from(data) : (new Uint8Array(data) as Buffer);
}

/** Hands the bytes to the user as a download: the browser counterpart of writing a file. */
export function downloadInBrowser(fileName: string, blob: Blob): void {
  const document = getWindow()?.document;
  const urls = (globalThis as { URL?: ObjectUrlFactory }).URL;
  if (!document || !urls) {
    throw new Error(`Cannot save "${fileName}": no document is available to download it from.`);
  }
  const objectUrl = urls.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = objectUrl;
  link.download = fileName;
  document.body?.appendChild(link);
  link.click();
  document.body?.removeChild(link);
  // Revoked late: some browsers abort the download if the URL dies before they start reading it.
  setTimeout(() => urls.revokeObjectURL(objectUrl), 10_000).unref?.();
}

let browserWarningPrinted = false;

/** Warns once, in development only, that a browser bundle ships the API key to end users. */
export function warnAboutBrowserApiKeyExposure(): void {
  if (browserWarningPrinted || !isBrowser() || !isDevelopment()) {
    return;
  }
  browserWarningPrinted = true;
  console.warn(
    '[mitosis-excel-templater-client] Running in a browser: your API key is shipped with the page and ' +
      'can be read by anyone using it. Your quota, and the templates stored on your account, are exposed.\n' +
      'For production, either call the API from your own backend, or generate files locally with the private ' +
      `@mitosis/mitosis-excel-templater package, which needs no API key in the browser: ${PRICING_URL}\n` +
      'This warning is only printed in development mode.',
  );
}
