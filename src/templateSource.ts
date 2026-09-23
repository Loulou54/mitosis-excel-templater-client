import { readFile } from './fileIo.js';
import type { ClientOptions, TemplateRef, TemplateSource } from './types.js';
import { request } from './http.js';

export const XLSX_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const DEFAULT_TEMPLATE_FILE_NAME = 'template.xlsx';

export function isTemplateRef(source: TemplateSource): source is TemplateRef {
  return (
    typeof source === 'object' &&
    source !== null &&
    'templateId' in source &&
    typeof (source as TemplateRef).templateId === 'string'
  );
}

function isBinary(source: TemplateSource): source is Buffer | ArrayBuffer | Uint8Array {
  return source instanceof ArrayBuffer || ArrayBuffer.isView(source);
}

function isHttpUrl(source: string): boolean {
  return /^https?:\/\//i.test(source);
}

/** Derives the file name sent to the API, used by the engine to name generated types. */
export function templateFileName(source: TemplateSource): string {
  if (typeof source === 'string') {
    const name = source.split(/[/\\]/).pop()?.split('?')[0];
    return name || DEFAULT_TEMPLATE_FILE_NAME;
  }
  return DEFAULT_TEMPLATE_FILE_NAME;
}

/**
 * Reads the template bytes from wherever they live: memory, the local file system, or an http(s) URL.
 * Stored templates (`{ templateId }`) are never read locally - the API resolves them server side.
 */
export async function readTemplateBytes(
  source: TemplateSource,
  options?: ClientOptions,
): Promise<Uint8Array> {
  if (isBinary(source)) {
    return source instanceof ArrayBuffer ? new Uint8Array(source) : new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  }
  if (typeof source === 'string') {
    if (!isHttpUrl(source)) {
      return readFile(source);
    }
    const response = await request({
      path: '',
      method: 'GET',
      options: { ...options, baseUrl: source },
      requiresApiKey: false,
    });
    return new Uint8Array(await response.arrayBuffer());
  }
  throw new TypeError(
    'Unsupported template source: expected a Buffer, an ArrayBuffer, a file path, an http(s) URL or { templateId }.',
  );
}

/** Appends either the `file` part or the `templateId` part expected by the API. */
export function appendTemplate(form: FormData, source: TemplateSource, bytes?: Uint8Array): void {
  if (isTemplateRef(source)) {
    form.append('templateId', source.templateId);
    return;
  }
  if (!bytes) {
    throw new TypeError('Template bytes are required for non stored templates.');
  }
  form.append('file', new Blob([bytes], { type: XLSX_MIME_TYPE }), templateFileName(source));
}
