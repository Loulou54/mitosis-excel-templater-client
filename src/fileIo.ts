import { downloadInBrowser, isBrowser } from './runtime.js';

type FsModule = typeof import('node:fs');

// Assembled at runtime so bundlers targeting the browser cannot statically resolve the Node
// built-in: webpack otherwise fails the build with `UnhandledSchemeError: node:fs`.
const FS_SPECIFIER = ['node', 'fs'].join(':');

let fsModule: Promise<FsModule> | undefined;

// Loaded lazily and dynamically so the module graph stays free of static Node built-in imports.
async function loadFs(): Promise<FsModule> {
  if (isBrowser()) {
    throw new Error('The file system is not available in the browser.');
  }
  if (!fsModule) {
    fsModule = import(/* webpackIgnore: true */ /* @vite-ignore */ FS_SPECIFIER) as Promise<FsModule>;
  }
  return fsModule;
}

/** Content types used when a generated file is downloaded instead of written to disk. */
function mimeTypeOf(path: string): string {
  if (path.endsWith('.ts')) {
    return 'text/plain; charset=utf-8';
  }
  if (path.endsWith('.json')) {
    return 'application/json; charset=utf-8';
  }
  return 'application/octet-stream';
}

export async function readFile(path: string): Promise<Buffer> {
  const fs = await loadFs();
  return fs.promises.readFile(path);
}

/** Writes to the file system on Node.js, and downloads the file in the browser. */
export async function writeFile(path: string, data: string | Uint8Array): Promise<void> {
  if (isBrowser()) {
    downloadInBrowser(path.split(/[/\\]/).pop() ?? path, new Blob([data], { type: mimeTypeOf(path) }));
    return;
  }
  const fs = await loadFs();
  await fs.promises.writeFile(path, data);
}

/** Last path segment without its extension, e.g. `src/types/CarsData.d.ts` -> `CarsData`. */
export function baseNameWithoutExtension(path: string): string {
  return path.split(/[/\\]/).pop()?.split('.')[0] ?? '';
}

/** Applies the engine's naming rule for generated type files: always end up with a single `.d.ts`. */
export function normalizeTypescriptPath(path: string): string {
  return path.replace('.d', '').replace('.ts', '') + '.d.ts';
}

/** Applies the engine's naming rule for generated schemas: always end up with a single `.schema.json`. */
export function normalizeJsonSchemaPath(path: string): string {
  return path.replace('.schema', '').replace('.json', '') + '.schema.json';
}
