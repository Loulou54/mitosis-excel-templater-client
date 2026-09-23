import { expect } from 'chai';
import { TemplatesClient } from './TemplatesClient.js';
import { getEngineVersion } from './version.js';
import { resetConfig } from './config.js';
import { binaryResponse, createFetchStub, formFileBytes, jsonResponse } from './testUtils.js';

const API_KEY = 'test-api-key';
const TEMPLATE_BYTES = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);

describe('TemplatesClient', () => {
  afterEach(() => resetConfig());

  it('unwraps the templates envelope', async () => {
    const templates = [
      { id: '1', templateId: 'invoice', fileName: 'invoice.xlsx', fileSize: 42, createdAt: '2026-01-01' },
    ];
    const stub = createFetchStub(jsonResponse({ templates }));

    const result = await new TemplatesClient({ apiKey: API_KEY, fetch: stub.fetch }).listTemplates();

    expect(stub.lastCall().url).to.equal('https://www.mitosis-excel-templater.dev/api/templates');
    expect(result).to.deep.equal(templates);
  });

  it('uploads a template with both the file and the templateId parts', async () => {
    const stub = createFetchStub(
      jsonResponse({ message: 'ok', templateId: 'invoice', fileName: 'template.xlsx', fileSize: 4 }),
    );

    await new TemplatesClient({ apiKey: API_KEY, fetch: stub.fetch }).uploadTemplate(
      'invoice',
      TEMPLATE_BYTES,
    );

    const call = stub.lastCall();
    expect(call.method).to.equal('PUT');
    expect(call.form!.get('templateId')).to.equal('invoice');
    expect(await formFileBytes(call.form!)).to.deep.equal(TEMPLATE_BYTES);
  });

  it('passes templateId as a query parameter when downloading', async () => {
    const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));

    const buffer = await new TemplatesClient({ apiKey: API_KEY, fetch: stub.fetch }).downloadTemplate(
      'invoice',
    );

    expect(stub.lastCall().url).to.equal(
      'https://www.mitosis-excel-templater.dev/api/template?templateId=invoice',
    );
    expect(new Uint8Array(buffer)).to.deep.equal(TEMPLATE_BYTES);
  });

  it('deletes a stored template', async () => {
    const stub = createFetchStub(jsonResponse({ message: 'Template deleted successfully' }));

    await new TemplatesClient({ apiKey: API_KEY, fetch: stub.fetch }).deleteTemplate('invoice');

    expect(stub.lastCall().method).to.equal('DELETE');
    expect(stub.lastCall().url).to.contain('templateId=invoice');
  });
});

describe('getEngineVersion', () => {
  afterEach(() => resetConfig());

  it('queries the public version endpoint without an api key', async () => {
    const stub = createFetchStub(
      jsonResponse({ name: '@mitosis/mitosis-excel-templater', version: '1.1.1' }),
    );

    const version = await getEngineVersion({ fetch: stub.fetch });

    expect(stub.lastCall().url).to.contain('/api/mitosis-excel-templater-version');
    expect(stub.lastCall().headers['x-api-key']).to.equal(undefined);
    expect(version.version).to.equal('1.1.1');
  });
});
