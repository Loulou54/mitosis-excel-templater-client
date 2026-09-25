/**
 * End to end tests running on Node.js against the real Mitosis Excel Templater API.
 *
 * They are skipped unless an API key is available, since they consume the account quota:
 *   MITOSIS_EXCEL_TEMPLATER_API_KEY=your-api-key npm run test:integration
 *
 * Target another deployment with MITOSIS_EXCEL_TEMPLATER_BASE_URL, e.g. http://localhost:3000.
 */
import { expect } from 'chai';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateRawSync } from 'node:zlib';
import { API_KEY_ENV_VAR } from './config.js';
import { AuthenticationError, MitosisApiError, TemplateNotFoundError } from './errors.js';
import { ExcelTemplater } from './ExcelTemplater.js';
import { TemplatesClient } from './TemplatesClient.js';
import type { ClientOptions, TemplateData } from './types.js';
import { getEngineVersion } from './version.js';

const apiKey = process.env[API_KEY_ENV_VAR];
const baseUrl = process.env.MITOSIS_EXCEL_TEMPLATER_BASE_URL;
const options: ClientOptions = { apiKey, ...(baseUrl ? { baseUrl } : {}) };

/** The whole suite is skipped rather than failed when no key is configured, so CI stays green. */
const describeApi = apiKey ? describe : describe.skip;

const TEMPLATE_ROWS = [
  ['{name}', '{surName}'],
  ['Job: {job}', '', '{?job}'],
  ['{brand}', '{model}', '{#carList}'],
  ['{color}', '{=year}', '{/carList}'],
];

const TEMPLATE_DATA: TemplateData = {
  name: 'Louis',
  surName: 'Durand',
  job: 'Software Engineer',
  carList: [
    { brand: 'Citroën', model: 'C4', color: 'Blue', year: '2011' },
    { brand: 'Peugeot', model: '3008', color: 'Brown', year: 2012 },
  ],
};

describeApi('Mitosis Excel Templater API (integration)', function () {
  // Cold starts on the hosted API are well above Mocha's default 2s budget.
  this.timeout(120_000);

  let template: Buffer;
  let workDir: string;

  before(() => {
    template = buildXlsxTemplate(TEMPLATE_ROWS);
    workDir = mkdtempSync(join(tmpdir(), 'mitosis-client-it-'));
  });

  after(() => {
    rmSync(workDir, { recursive: true, force: true });
  });

  describe('getEngineVersion', () => {
    it('reports the engine powering the API without any key', async () => {
      const version = await getEngineVersion(baseUrl ? { baseUrl } : undefined);

      expect(version.name).to.be.a('string').and.not.be.empty;
      expect(version.version).to.match(/^\d+\.\d+\.\d+/);
    });
  });

  describe('authentication', () => {
    it('rejects an unknown api key with an AuthenticationError', async () => {
      const templater = new ExcelTemplater(template, { ...options, apiKey: 'not-a-valid-api-key' });

      const error = await templater.saveAsExcel(TEMPLATE_DATA).catch((e: unknown) => e);

      expect(error).to.be.instanceOf(AuthenticationError);
      expect((error as AuthenticationError).status).to.equal(401);
    });
  });

  describe('saveAsExcel', () => {
    it('fills the template tags and returns a real xlsx workbook', async () => {
      const templater = new ExcelTemplater(template, options);

      const workbook = await templater.saveAsExcel(TEMPLATE_DATA);

      expect(Buffer.isBuffer(workbook)).to.equal(true);
      expect(workbook.subarray(0, 2).toString('latin1')).to.equal('PK');

      const text = sheetText(workbook);
      expect(text).to.include('Louis').and.to.include('Durand');
      expect(text).to.include('Software Engineer');
      expect(text).to.include('Citroën').and.to.include('3008');
      expect(text).to.not.include('{name}');
      expect(text).to.not.include('{#carList}');
    });

    it('writes the workbook to disk when a file name is given', async () => {
      const target = join(workDir, 'generated.xlsx');
      const templater = new ExcelTemplater(template, options);

      const workbook = await templater.saveAsExcel(TEMPLATE_DATA, target);

      expect(existsSync(target)).to.equal(true);
      expect(readFileSync(target).equals(workbook)).to.equal(true);
    });

    it('reads the template from the local file system when given a path', async () => {
      const templatePath = join(workDir, 'cars.xlsx');
      writeFileSync(templatePath, template);

      const workbook = await new ExcelTemplater(templatePath, options).saveAsExcel(TEMPLATE_DATA);

      expect(sheetText(workbook)).to.include('Louis');
    });

    it('reports a malformed template as a MitosisApiError', async () => {
      const broken = buildXlsxTemplate([['{brand}', '{#carList}']]); // section opened, never closed
      const templater = new ExcelTemplater(broken, options);

      const error = await templater.saveAsExcel(TEMPLATE_DATA).catch((e: unknown) => e);

      expect(error).to.be.instanceOf(MitosisApiError);
      expect((error as MitosisApiError).status).to.be.greaterThan(399);
    });
  });

  describe('template introspection', () => {
    const templater = () => new ExcelTemplater(template, options);

    it('generates sample data covering every tag of the template', async () => {
      const sample = await templater().generateSampleData();

      expect(Object.keys(sample)).to.include.members(['name', 'surName', 'job', 'carList']);
      const carList = sample.carList as TemplateData[];
      expect(carList).to.be.an('array').with.lengthOf(1);
      expect(Object.keys(carList[0])).to.include.members(['brand', 'model', 'color', 'year']);
    });

    it('generates a JSON Schema describing the template data', async () => {
      const schema = (await templater().generateTemplateDataJsonSchema()) as {
        properties?: Record<string, unknown>;
      };

      expect(Object.keys(schema.properties ?? {})).to.include.members([
        'name',
        'surName',
        'job',
        'carList',
      ]);
    });

    it('marks every property as optional when asked to', async () => {
      const schema = (await templater().generateTemplateDataJsonSchema(undefined, true)) as {
        required?: string[];
      };

      expect(schema.required ?? []).to.be.empty;
    });

    it('generates a typescript type and saves it next to the caller code', async () => {
      const generated = await templater().generateTemplateDataTypescriptFile(
        join(workDir, 'CarsData.ts'),
      );

      expect(generated).to.include('export type CarsData =');
      expect(generated).to.include('carList');
      // The engine imports from the private package; the client must point users to this one.
      expect(generated).to.not.include("'@mitosis/mitosis-excel-templater'");

      const target = join(workDir, 'CarsData.d.ts');
      expect(existsSync(target)).to.equal(true);
      expect(readFileSync(target, 'utf8')).to.equal(generated);
    });
  });

  describe('TemplatesClient', () => {
    const templateId = `client-integration-${Date.now()}`;
    const client = new TemplatesClient(options);

    before(async () => {
      const result = await client.uploadTemplate(templateId, template);
      expect(result.templateId).to.equal(templateId);
      expect(result.fileSize).to.equal(template.length);
    });

    after(async () => {
      await client.deleteTemplate(templateId).catch(() => undefined);
    });

    it('lists the uploaded template', async () => {
      const templates = await client.listTemplates();

      const stored = templates.find((it) => it.templateId === templateId);
      expect(stored, `template ${templateId} missing from ${templates.length} listed`).to.exist;
      expect(stored!.fileSize).to.equal(template.length);
      expect(Date.parse(stored!.createdAt)).to.be.a('number').and.not.be.NaN;
    });

    it('downloads the stored template byte for byte', async () => {
      const downloaded = await client.downloadTemplate(templateId);

      expect(downloaded.equals(template)).to.equal(true);
    });

    it('downloads the stored template to a file', async () => {
      const target = join(workDir, 'downloaded.xlsx');

      await client.downloadTemplate(templateId, target);

      expect(readFileSync(target).equals(template)).to.equal(true);
    });

    it('generates a workbook from the stored template, without re-uploading it', async () => {
      const workbook = await new ExcelTemplater({ templateId }, options).saveAsExcel(TEMPLATE_DATA);

      expect(sheetText(workbook)).to.include('Louis').and.to.include('Citroën');
    });

    it('fails with a TemplateNotFoundError on an unknown template id', async () => {
      const error = await client.downloadTemplate('does-not-exist-ever').catch((e: unknown) => e);

      expect(error).to.be.instanceOf(TemplateNotFoundError);
      expect((error as TemplateNotFoundError).status).to.equal(404);
    });

    it('deletes the stored template', async () => {
      await client.deleteTemplate(templateId);

      const error = await client.downloadTemplate(templateId).catch((e: unknown) => e);
      expect(error).to.be.instanceOf(TemplateNotFoundError);
    });
  });
});

// --- Minimal xlsx tooling -----------------------------------------------------------------------
// Building and reading the workbooks here keeps the package dependency-free: adding ExcelJS just for
// the tests would hide the fact that the client itself never needs an Excel engine.

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let crc = i;
    for (let bit = 0; bit < 8; bit++) {
      crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table[i] = crc >>> 0;
  }
  return table;
})();

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** Writes an uncompressed zip archive, the container format of every xlsx file. */
function zip(files: Array<{ name: string; content: string }>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;

  for (const file of files) {
    const name = Buffer.from(file.name, 'utf8');
    const data = Buffer.from(file.content, 'utf8');
    const checksum = crc32(data);

    const local = Buffer.alloc(30 + name.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 8); // stored, no compression
    local.writeUInt16LE(0x21, 12); // 1980-01-01, so archives are byte stable
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    name.copy(local, 30);

    const central = Buffer.alloc(46 + name.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    name.copy(central, 46);

    locals.push(local, data);
    centrals.push(central);
    offset += local.length + data.length;
  }

  const directory = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);

  return Buffer.concat([...locals, directory, end]);
}

/** Reads a zip archive entry by entry. Sizes come from the central directory, as streamed
 *  archives - like the ones the API returns - leave them empty in the local headers. */
function unzip(archive: Buffer): Map<string, Buffer> {
  let end = -1;
  for (let i = archive.length - 22; i >= 0; i--) {
    if (archive.readUInt32LE(i) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) {
    throw new Error('Not a zip archive: end of central directory not found');
  }

  const entries = new Map<string, Buffer>();
  const count = archive.readUInt16LE(end + 10);
  let pointer = archive.readUInt32LE(end + 16);

  for (let i = 0; i < count; i++) {
    const method = archive.readUInt16LE(pointer + 10);
    const compressedSize = archive.readUInt32LE(pointer + 20);
    const nameLength = archive.readUInt16LE(pointer + 28);
    const extraLength = archive.readUInt16LE(pointer + 30);
    const commentLength = archive.readUInt16LE(pointer + 32);
    const localOffset = archive.readUInt32LE(pointer + 42);
    const name = archive.toString('utf8', pointer + 46, pointer + 46 + nameLength).replace(/\\/g, '/');

    const dataStart =
      localOffset + 30 + archive.readUInt16LE(localOffset + 26) + archive.readUInt16LE(localOffset + 28);
    const raw = archive.subarray(dataStart, dataStart + compressedSize);
    entries.set(name, method === 0 ? Buffer.from(raw) : inflateRawSync(raw));

    pointer += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/** Every string a generated workbook displays, whether stored inline or in the shared string table. */
function sheetText(workbook: Buffer): string {
  const entries = unzip(workbook);
  return [...entries]
    .filter(([name]) => name === 'xl/sharedStrings.xml' || name.startsWith('xl/worksheets/'))
    .map(([, content]) => content.toString('utf8'))
    .join('\n');
}

function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 0 -> A, 25 -> Z, 26 -> AA */
function columnName(index: number): string {
  let name = '';
  for (let n = index; n >= 0; n = Math.floor(n / 26) - 1) {
    name = String.fromCharCode(65 + (n % 26)) + name;
  }
  return name;
}

/** Builds a valid xlsx workbook whose first sheet holds the given rows of text. */
function buildXlsxTemplate(rows: string[][]): Buffer {
  const strings: string[] = [];
  const stringIndex = (text: string): number => {
    const existing = strings.indexOf(text);
    return existing >= 0 ? existing : strings.push(text) - 1;
  };

  const sheetData = rows
    .map((cells, rowIndex) => {
      const row = rowIndex + 1;
      const xml = cells
        .map((text, column) =>
          text
            ? `<c r="${columnName(column)}${row}" t="s"><v>${stringIndex(text)}</v></c>`
            : '',
        )
        .join('');
      return `<row r="${row}">${xml}</row>`;
    })
    .join('');

  const sharedStrings = strings
    .map((text) => `<si><t xml:space="preserve">${escapeXml(text)}</t></si>`)
    .join('');

  const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  return zip([
    {
      name: '[Content_Types].xml',
      content:
        `${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>',
    },
    {
      name: '_rels/.rels',
      content:
        `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${relationships}/officeDocument" Target="xl/workbook.xml"/>` +
        '</Relationships>',
    },
    {
      name: 'xl/workbook.xml',
      content:
        `${declaration}<workbook xmlns="${main}" xmlns:r="${relationships}">` +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      content:
        `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${relationships}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${relationships}/sharedStrings" Target="sharedStrings.xml"/>` +
        `<Relationship Id="rId3" Type="${relationships}/styles" Target="styles.xml"/>` +
        '</Relationships>',
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      content: `${declaration}<worksheet xmlns="${main}"><sheetData>${sheetData}</sheetData></worksheet>`,
    },
    {
      name: 'xl/sharedStrings.xml',
      content:
        `${declaration}<sst xmlns="${main}" count="${strings.length}" uniqueCount="${strings.length}">` +
        `${sharedStrings}</sst>`,
    },
    {
      name: 'xl/styles.xml',
      content:
        `${declaration}<styleSheet xmlns="${main}">` +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
        '</styleSheet>',
    },
  ]);
}
