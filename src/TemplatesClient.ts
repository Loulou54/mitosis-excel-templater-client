import { writeFile } from './fileIo.js';
import { requestBuffer, requestJson } from './http.js';
import { appendTemplate, readTemplateBytes } from './templateSource.js';
import type {
  ClientOptions,
  TemplateSource,
  TemplateSummary,
  UploadTemplateResult,
} from './types.js';

/**
 * Manages the Excel templates stored on your account.
 *
 * Uploading a template once and referencing it later by id avoids re-sending the same file on every
 * generation: `new ExcelTemplater({ templateId: 'invoice' })`.
 */
export class TemplatesClient {
  private readonly options: ClientOptions;

  constructor(options: ClientOptions = {}) {
    this.options = options;
  }

  /** Lists every template stored on the account. */
  public async listTemplates(): Promise<TemplateSummary[]> {
    const { templates } = await requestJson<{ templates: TemplateSummary[] }>({
      path: '/api/templates',
      method: 'GET',
      options: this.options,
    });
    return templates;
  }

  /**
   * Stores a template under the given id, replacing it if the id already exists.
   * @param templateId the identifier you will reference the template with.
   * @param template the template as a Buffer, a local file path or an http(s) URL.
   */
  public async uploadTemplate(
    templateId: string,
    template: Exclude<TemplateSource, { templateId: string }>,
  ): Promise<UploadTemplateResult> {
    const form = new FormData();
    appendTemplate(form, template, await readTemplateBytes(template, this.options));
    form.append('templateId', templateId);

    return requestJson<UploadTemplateResult>({
      path: '/api/template',
      method: 'PUT',
      body: form,
      options: this.options,
    });
  }

  /**
   * Downloads a stored template.
   * @param templateId the identifier of the stored template.
   * @param fileName (optional) the path under which the template will be saved.
   */
  public async downloadTemplate(templateId: string, fileName?: string): Promise<Buffer> {
    const buffer = await requestBuffer({
      path: '/api/template',
      method: 'GET',
      query: { templateId },
      options: this.options,
    });
    if (fileName) {
      await writeFile(fileName, buffer);
    }
    return buffer;
  }

  /** Permanently deletes a stored template. */
  public async deleteTemplate(templateId: string): Promise<void> {
    await requestJson<{ message: string }>({
      path: '/api/template',
      method: 'DELETE',
      query: { templateId },
      options: this.options,
    });
  }
}
