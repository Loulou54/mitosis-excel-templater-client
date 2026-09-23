type FsModule = typeof import('node:fs');

let fsModule: Promise<FsModule> | undefined;

// Loaded lazily and dynamically so the module graph stays free of static Node built-in imports.
async function loadFs(): Promise<FsModule> {
  if (!fsModule) {
    fsModule = import('node:fs');
  }
  return fsModule;
}

export async function readFile(path: string): Promise<Buffer> {
  const fs = await loadFs();
  return fs.promises.readFile(path);
}

export async function writeFile(path: string, data: string | Uint8Array): Promise<void> {
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
