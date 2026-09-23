import { expect } from 'chai';
import { ExcelTemplater } from './ExcelTemplater.js';
import { configure, resetConfig, API_KEY_ENV_VAR } from './config.js';
import {
  AuthenticationError,
  MitosisApiError,
  RateLimitError,
  SubscriptionError,
  TemplateNotFoundError,
  TemplateSyntaxError,
} from './errors.js';
import {
  binaryResponse,
  createFetchStub,
  formFileBytes,
  jsonResponse,
  textResponse,
} from './testUtils.js';

const API_KEY = 'test-api-key';
const TEMPLATE_BYTES = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);

describe('ExcelTemplater', () => {
  afterEach(() => {
    resetConfig();
    delete process.env[API_KEY_ENV_VAR];
  });

  describe('authentication', () => {
    it('sends the api key as the x-api-key header', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      await templater.saveAsExcel({ name: 'Louis' });

      expect(stub.lastCall().headers['x-api-key']).to.equal(API_KEY);
    });

    it('prefers the constructor key over configure() and the environment', async () => {
      configure({ apiKey: 'global-key' });
      process.env[API_KEY_ENV_VAR] = 'env-key';
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));

      await new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch }).saveAsExcel({});

      expect(stub.lastCall().headers['x-api-key']).to.equal(API_KEY);
    });

    it('falls back to configure() when no key is passed', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      configure({ apiKey: 'global-key', fetch: stub.fetch });

      await new ExcelTemplater(TEMPLATE_BYTES).saveAsExcel({});

      expect(stub.lastCall().headers['x-api-key']).to.equal('global-key');
    });

    it('falls back to the environment variable last', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      process.env[API_KEY_ENV_VAR] = 'env-key';

      await new ExcelTemplater(TEMPLATE_BYTES, { fetch: stub.fetch }).saveAsExcel({});

      expect(stub.lastCall().headers['x-api-key']).to.equal('env-key');
    });

    it('throws before any request when no key is available', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { fetch: stub.fetch });

      const error = await templater.saveAsExcel({}).catch((e: unknown) => e);

      expect(error).to.be.instanceOf(AuthenticationError);
      expect(stub.calls).to.have.length(0);
    });
  });

  describe('saveAsExcel', () => {
    it('posts the template as a file part and the data as a JSON string', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      await templater.saveAsExcel({ name: 'Louis', carList: [{ model: 'C4' }] });

      const call = stub.lastCall();
      expect(call.url).to.equal('https://www.mitosis-excel-templater.dev/api/generate');
      expect(call.method).to.equal('POST');
      expect(await formFileBytes(call.form!)).to.deep.equal(TEMPLATE_BYTES);
      expect(JSON.parse(call.form!.get('data') as string)).to.deep.equal({
        name: 'Louis',
        carList: [{ model: 'C4' }],
      });
    });

    it('sends templateId instead of a file for stored templates', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      const templater = new ExcelTemplater({ templateId: 'invoice' }, {
        apiKey: API_KEY,
        fetch: stub.fetch,
      });

      await templater.saveAsExcel({});

      const form = stub.lastCall().form!;
      expect(form.get('templateId')).to.equal('invoice');
      expect(form.get('file')).to.equal(null);
    });

    it('returns the generated workbook as a Buffer', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const buffer = await templater.saveAsExcel({});

      expect(Buffer.isBuffer(buffer)).to.equal(true);
      expect(new Uint8Array(buffer)).to.deep.equal(TEMPLATE_BYTES);
    });

    it('honours a custom baseUrl', async () => {
      const stub = createFetchStub(binaryResponse(TEMPLATE_BYTES));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, {
        apiKey: API_KEY,
        baseUrl: 'http://localhost:3000/',
        fetch: stub.fetch,
      });

      await templater.saveAsExcel({});

      expect(stub.lastCall().url).to.equal('http://localhost:3000/api/generate');
    });
  });

  describe('generateSampleData', () => {
    it('unwraps the sampleData envelope', async () => {
      const stub = createFetchStub(jsonResponse({ sampleData: { name: 'name', year: 0 } }));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const data = await templater.generateSampleData();

      expect(stub.lastCall().url).to.contain('/api/generate-sample-data');
      expect(data).to.deep.equal({ name: 'name', year: 0 });
    });
  });

  describe('generateTemplateDataJsonSchema', () => {
    it('unwraps the schema envelope and forwards propsAreOptional', async () => {
      const stub = createFetchStub(jsonResponse({ schema: { type: 'object' }, propsAreOptional: true }));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const schema = await templater.generateTemplateDataJsonSchema(undefined, true);

      expect(stub.lastCall().form!.get('propsAreOptional')).to.equal('true');
      expect(schema).to.deep.equal({ type: 'object' });
    });
  });

  describe('generateTemplateDataTypescriptFile', () => {
    const generated =
      "import { CellValue, Section } from '@mitosis/mitosis-excel-templater';\n\n" +
      'export type Cars = { name: CellValue };';

    it('rewrites the engine import to this package', async () => {
      const stub = createFetchStub(textResponse(generated));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const result = await templater.generateTemplateDataTypescriptFile();

      expect(result).to.contain("from 'mitosis-excel-templater-client'");
      expect(result).to.not.contain('@mitosis/mitosis-excel-templater');
    });

    it('rewrites double quoted imports too', async () => {
      const stub = createFetchStub(
        textResponse('import { CellValue } from "@mitosis/mitosis-excel-templater";'),
      );
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const result = await templater.generateTemplateDataTypescriptFile();

      expect(result).to.contain('from "mitosis-excel-templater-client"');
    });
  });

  describe('error mapping', () => {
    const cases: [number, unknown, unknown][] = [
      [401, { error: 'Invalid API key' }, AuthenticationError],
      [403, { error: 'npm package access requires subscription' }, SubscriptionError],
      [404, { error: 'Template not found' }, TemplateNotFoundError],
      [429, { error: 'Rate limit exceeded', remaining: 0 }, RateLimitError],
      [500, { error: 'Internal server error' }, MitosisApiError],
    ];

    cases.forEach(([status, body, expected]) => {
      it(`maps HTTP ${status} to ${(expected as { name: string }).name}`, async () => {
        const stub = createFetchStub(jsonResponse(body, status));
        const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

        const error = await templater.saveAsExcel({}).catch((e: unknown) => e);

        expect(error).to.be.instanceOf(expected as never);
        expect((error as MitosisApiError).status).to.equal(status);
        expect((error as MitosisApiError).message).to.equal((body as { error: string }).error);
      });
    });

    it('exposes the remaining quota on a rate limit error', async () => {
      const stub = createFetchStub(jsonResponse({ error: 'Rate limit exceeded', remaining: 0 }, 429));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const error = (await templater.saveAsExcel({}).catch((e: unknown) => e)) as RateLimitError;

      expect(error.remaining).to.equal(0);
    });

    it('maps template formatting errors to TemplateSyntaxError', async () => {
      const message = 'Excel template formatting error in worksheet "Sheet1": section "items" was not closed.';
      const stub = createFetchStub(jsonResponse({ error: message }, 500));
      const templater = new ExcelTemplater(TEMPLATE_BYTES, { apiKey: API_KEY, fetch: stub.fetch });

      const error = await templater.saveAsExcel({}).catch((e: unknown) => e);

      expect(error).to.be.instanceOf(TemplateSyntaxError);
      expect((error as TemplateSyntaxError).message).to.equal(message);
    });
  });

  describe('removeEmptyObjects', () => {
    it('removes nil values and empty objects recursively', () => {
      const data = {
        name: 'Louis',
        company: undefined,
        address: {},
        nested: { empty: {}, kept: 1 },
      };

      ExcelTemplater.removeEmptyObjects(data);

      expect(data).to.deep.equal({ name: 'Louis', nested: { kept: 1 } });
    });

    it('keeps Date values untouched', () => {
      const date = new Date('2023-12-25');
      const data = { genDate: date };

      ExcelTemplater.removeEmptyObjects(data);

      expect(data.genDate).to.equal(date);
    });
  });

  describe('formatDates', () => {
    it('converts ISO date strings into Date objects, recursively', () => {
      const data = {
        invoiceDate: '2023-12-25',
        label: 'not a date',
        items: [{ orderDate: '2023-12-20' }],
      };

      ExcelTemplater.formatDates(data);

      expect(data.invoiceDate).to.be.instanceOf(Date);
      expect(data.label).to.equal('not a date');
      expect((data.items[0] as { orderDate: unknown }).orderDate).to.be.instanceOf(Date);
    });
  });
});
