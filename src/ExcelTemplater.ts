import {
  baseNameWithoutExtension,
  normalizeJsonSchemaPath,
  normalizeTypescriptPath,
  writeFile,
} from './fileIo.js';
import { request, requestJson, requestText } from './http.js';
import { toBuffer } from './runtime.js';
import { appendTemplate, isTemplateRef, readTemplateBytes } from './templateSource.js';
import type { ClientOptions, TemplateData, TemplateRef, TemplateSource } from './types.js';

const CORE_PACKAGE_NAME = '@mitosis/mitosis-excel-templater';
const CLIENT_PACKAGE_NAME = 'mitosis-excel-templater-client';

/** Mirrors the engine's `toAlphanumCamelCase` so generated type names match. */
function toAlphanumCamelCase(value: string): string {
  return value
    .replace(/[^A-Z0-9]+/gi, '_')
    .split('_')
    .map((it) => it.charAt(0).toUpperCase() + it.substring(1))
    .join('');
}

function isNil(value: unknown): value is null | undefined {
  return value === null || value === undefined;
}

/**
 * Inspiring from Mustache syntax, ExcelTemplater creates an Excel file by feeding structured labelled
 * data into an Excel template containing replaceable tags.
 *
 * This client exposes the same interface as the `@mitosis/mitosis-excel-templater` engine, but every
 * generation runs on the hosted Mitosis Excel Templater API instead of locally.
 *
 * - Replaceable tags have the following format in the Excel template: {tagName}
 * - You can repeat sections by annotating the start row with {#sectionName} and the end row with
 *   {/sectionName}. If empty, the section is not shown. The data under "sectionName" must be an array
 *   of objects, or an object directly.
 * - You can display a row conditionally with {?tagName}. If the value under tagName is undefined,
 *   false or an empty array, the line is not displayed.
 * - You can force a field to be evaluated as a number with {=tagName}. The number format chosen for
 *   the cell in the template is applied.
 *
 * Example:
 * > Excel Template:
 * | {name}    | {surName}  |
 * | {job}     | {?job}     |
 * | {carName} | {color}    | {#carList} |
 * | {brand}   | {=year}    | {/carList} |
 *
 * > Data:
 * {
 *  "name": "Louis",
 *  "surName": "Durand",
 *  "carList": [
 *    { "carName": "C4", "color": "Blue", "brand": "Citroën", "year": "2011" },
 *    { "carName": "3008", "color": "Brown", "brand": "Peugeot", "year": 2011 }
 *  ]
 * }
 */
export class ExcelTemplater {
  /**
   * Utility function to help recursively remove empty fields within your template data structure.
   * @param object TemplateData object with empty fields to cleanup.
   */
  public static removeEmptyObjects(object: TemplateData): void {
    Object.keys(object).forEach((key) => {
      const value = object[key];
      if (isNil(value)) {
        delete object[key];
      } else if (typeof value === 'object' && !(value instanceof Date)) {
        this.removeEmptyObjects(value as TemplateData);
        if (Object.keys(value).length === 0) {
          delete object[key];
        }
      }
    });
  }

  /**
   * Utility function to help recursively transform date strings into javascript Date objects within
   * your template data structure.
   * @param object TemplateData object with date fields formatted as strings.
   */
  public static formatDates(object: TemplateData): void {
    Object.keys(object).forEach((key) => {
      const value = object[key];
      if (typeof value === 'string' && /\d\d\d\d-\d\d-\d\d/.test(value)) {
        object[key] = new Date(value);
      } else if (!isNil(value) && typeof value === 'object' && !(value instanceof Date)) {
        this.formatDates(value as TemplateData);
      }
    });
  }

  private readonly source: TemplateSource;
  private readonly options: ClientOptions;
  private templateBytes?: Promise<Uint8Array>;

  /**
   * Creates an ExcelTemplater for the Excel template file given as Buffer.
   * @param excelFileAsBuffer the Buffer containing the raw Excel template file.
   * @param options (optional) API key and other client options.
   */
  constructor(excelFileAsBuffer: Buffer | ArrayBuffer | Uint8Array, options?: ClientOptions);

  /**
   * Creates an ExcelTemplater for the Excel template file located at the given path or URL.
   * @param templateFileToFetch a local file path on Node.js, or a URL the template is fetched from.
   *   In the browser there is no file system, so the string is always fetched, relative to the page.
   * @param options (optional) API key and other client options.
   */
  constructor(templateFileToFetch: string, options?: ClientOptions);

  /**
   * Creates an ExcelTemplater for a template already stored on your account. The template is never
   * uploaded again, which makes repeated generations noticeably lighter.
   * @param storedTemplate the identifier of a template uploaded through TemplatesClient.
   * @param options (optional) API key and other client options.
   */
  constructor(storedTemplate: TemplateRef, options?: ClientOptions);

  constructor(source: TemplateSource, options: ClientOptions = {}) {
    this.source = source;
    this.options = options;
  }

  /**
   * Generates the typescript type corresponding to the fields found in the given Excel template file.
   * You can use this generated file in your source code for type-checking your input TemplateData!
   * @param saveResultToFile (optional) the path to the .ts file where you want to save the generated type.
   *   (As a file in the file system for Node.js, or as a download in the browser.)
   * @param propsAreOptional (optional, default: false) true to set all properties on the generated type as optional.
   * @returns a string containing the generated type.
   */
  public async generateTemplateDataTypescriptFile(
    saveResultToFile?: string,
    propsAreOptional = false,
  ): Promise<string> {
    const form = await this.buildForm(propsAreOptional);
    const generated = await requestText({
      path: '/api/generate-template-data-typescript-file',
      method: 'POST',
      body: form,
      options: this.options,
    });

    let result = this.rewriteImportSpecifier(generated);
    if (saveResultToFile) {
      result = renameExportedType(result, toAlphanumCamelCase(baseNameWithoutExtension(saveResultToFile)));
      const target = normalizeTypescriptPath(saveResultToFile);
      await writeFile(target, result);
      console.log('=> Successfully saved generated type to ' + target);
    }
    return result;
  }

  /**
   * Generates a JSON Schema (Draft 7) corresponding to the fields found in the given Excel template file.
   * @param saveResultToFile (optional) the path to the .schema.json file where you want to save the generated schema.
   *   (As a file in the file system for Node.js, or as a download in the browser.)
   * @param propsAreOptional (optional, default: false) true to set all properties on the generated schema as optional.
   * @returns an object containing the generated JSON Schema.
   */
  public async generateTemplateDataJsonSchema(
    saveResultToFile?: string,
    propsAreOptional = false,
  ): Promise<object> {
    const form = await this.buildForm(propsAreOptional);
    const schema = await requestJson<object>({
      path: '/api/generate-template-data-json-schema',
      method: 'POST',
      body: form,
      options: this.options,
    });

    if (saveResultToFile) {
      const target = normalizeJsonSchemaPath(saveResultToFile);
      await writeFile(target, JSON.stringify(schema, null, 2));
      console.log('=> Successfully saved generated schema to ' + target);
    }
    return schema;
  }

  /**
   * Generates a sample data structure corresponding to the fields found in the given Excel template file.
   * Strings are set with their key as value, numbers with 0, dates with the current date and booleans
   * with true. Sections are arrays with one sample element.
   * @returns a TemplateData structure containing sample data.
   */
  public async generateSampleData(): Promise<TemplateData> {
    const form = await this.buildForm();
    return requestJson<TemplateData>({
      path: '/api/generate-sample-data',
      method: 'POST',
      body: form,
      options: this.options,
    });
  }

  /**
   * Generates an Excel file based on the template, populated with the templateData, and saves it under fileName.
   * @param templateData data structure containing the data to insert in the given Excel template.
   * @param fileName (optional) the path or file name under which the generated Excel file will be saved.
   *   (As a file in the file system for Node.js, or as a download in the browser.)
   * @return a Buffer containing the generated Excel file, or a Uint8Array in the browser.
   */
  public async saveAsExcel(templateData: TemplateData, fileName?: string): Promise<Buffer> {
    const form = await this.buildForm();
    form.append('data', JSON.stringify(templateData));

    const response = await request({
      path: '/api/generate',
      method: 'POST',
      body: form,
      options: this.options,
    });
    const buffer = toBuffer(await response.arrayBuffer());

    if (fileName) {
      await writeFile(fileName, buffer);
    }
    return buffer;
  }

  private async buildForm(propsAreOptional?: boolean): Promise<FormData> {
    const form = new FormData();
    appendTemplate(form, this.source, await this.loadTemplateBytes());
    if (propsAreOptional !== undefined) {
      form.append('propsAreOptional', String(propsAreOptional));
    }
    return form;
  }

  private async loadTemplateBytes(): Promise<Uint8Array | undefined> {
    if (isTemplateRef(this.source)) {
      return undefined;
    }
    if (!this.templateBytes) {
      this.templateBytes = readTemplateBytes(this.source, this.options);
    }
    return this.templateBytes;
  }

  /** The engine emits an import from the private package; public users need this package instead. */
  private rewriteImportSpecifier(generated: string): string {
    const pattern = new RegExp(`(['"])${CORE_PACKAGE_NAME.replace('/', '\\/')}\\1`, 'g');
    return generated.replace(pattern, `$1${CLIENT_PACKAGE_NAME}$1`);
  }
}

/** Keeps parity with the engine, which names the type after the output file rather than the template. */
function renameExportedType(generated: string, typeName: string): string {
  if (!typeName) {
    return generated;
  }
  return generated.replace(/export type\s+\w+\s*=/, `export type ${typeName} =`);
}
