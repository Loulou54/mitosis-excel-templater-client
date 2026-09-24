# mitosis-excel-templater-client

[![npm version](https://img.shields.io/npm/v/mitosis-excel-templater-client.svg)](https://www.npmjs.com/package/mitosis-excel-templater-client)
[![node](https://img.shields.io/node/v/mitosis-excel-templater-client.svg)](https://nodejs.org)
[![license](https://img.shields.io/npm/l/mitosis-excel-templater-client.svg)](./LICENSE)

Official Node.js / TypeScript client for the [Mitosis Excel Templater](https://www.mitosis-excel-templater.dev) API.

## What it is

Design your spreadsheet **once** in Excel — fonts, colours, merged cells, formulas, conditional formatting, images — and sprinkle it with mustache-style tags such as `{clientName}` or `{#invoiceLines}`. Then feed it structured data and get a fully formatted `.xlsx` back. Your layout stays where designers can edit it, and your code only ever deals with plain objects.

This package talks to the hosted Mitosis Excel Templater API, so the heavy Excel engine runs server-side:

- **Zero runtime dependencies** — no ExcelJS, no XML parsing, no multi-megabyte install.
- **Fully typed** — ships its own TypeScript declarations, plus tooling to generate a type for each of your templates.
- **Same interface as the engine** — identical to the `@mitosis/mitosis-excel-templater` package, so moving between hosted and offline generation is a one-line import change.
- **Dual CJS / ESM build**, Node.js 18 or later.
- **Runs in the browser too** — but it ships your API key with your bundle, see [Browser usage](#browser-usage).

## Live demo

You can try the templater **without installing anything**: upload a template, paste your data and download the result straight from the browser.

👉 **[Open the live demo](https://www.mitosis-excel-templater.dev)**

👉 **[Open the API playground](https://www.mitosis-excel-templater.dev/docs)** — it pre-fills your own API key once you are signed in.

## Installation

```bash
npm install mitosis-excel-templater-client
```

## Get an API key

1. **Create a free account** at [mitosis-excel-templater.dev/signup](https://www.mitosis-excel-templater.dev/signup).
2. **Open your dashboard** at [mitosis-excel-templater.dev/dashboard](https://www.mitosis-excel-templater.dev/dashboard).
3. **Copy your API key** displayed there. You can rotate it at any time with the *Regenerate* button — the previous key stops working immediately.
4. **Store it as an environment variable** named `MITOSIS_EXCEL_TEMPLATER_API_KEY`. The client picks it up automatically, so you never have to pass it explicitly:

   ```bash
   # .env, CI secret, container secret...
   MITOSIS_EXCEL_TEMPLATER_API_KEY=your-api-key
   ```

5. **Pick a plan** on the [pricing page](https://www.mitosis-excel-templater.dev/pricing) when the free tier gets tight. The free plan works out of the box, no card required.

> ⚠️ **Keep your API key server-side.** It grants access to your quota and to the templates stored on your account. Never commit it, and avoid shipping it inside a browser bundle or a mobile app. If a key leaks, regenerate it from the dashboard. The client *does* run in the browser, but at that cost — read [Browser usage](#browser-usage) before doing so.

## Browser usage

This client is **browser compatible**: it only uses `fetch`, `FormData` and `Blob`, so it works in any modern browser and bundler, with no Node.js polyfill.

> 🔐 **But calling the API from the browser exposes your API key.** Anything the browser sends, a user can read: your key ends up in the network tab and in your bundle, and with it anyone can burn your quota and read, overwrite or delete the templates stored on your account. There is no safe way around it — a hosted API call needs a key, and a key in a browser is a public key.

So, for browser applications, pick one of these instead:

1. **Generate locally with the private npm package** *(recommended)* — `@mitosis/mitosis-excel-templater` embeds the engine, runs entirely in the browser and needs **no API key at all**, since nothing leaves the page. Same interface, one-line import change. Subscribe on the [pricing page](https://www.mitosis-excel-templater.dev/pricing) and see [Offline generation](#offline-generation-without-any-api-call).
2. **Call the API from your own backend** — keep this client on your server, and expose your own endpoint to the browser.

Using this client in the browser is fine for a prototype, an internal tool behind your own network, or a key you are happy to rotate often. In that case, note that:

- the API key must be passed explicitly (`new ExcelTemplater(source, { apiKey })` or `configure({ apiKey })`), as there is no environment variable to read;
- there is no file system: a template given as a string is **fetched** as a URL, relative to the page (`new ExcelTemplater('templates/cars.xlsx')` fetches `/templates/cars.xlsx`), and the `fileName` argument of `saveAsExcel` triggers a **download** instead of writing to disk;
- `saveAsExcel` resolves with a `Uint8Array` instead of a `Buffer`, ready to be wrapped in a `Blob` for download;
- a warning about the key exposure is printed to the console **in development mode only**, never in a production build.

```ts
const bytes = await new ExcelTemplater({ templateId: 'cars' }, { apiKey }).saveAsExcel(templateData);
const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
```

## Quick start

Take an Excel template, `templates/cars.xlsx`:

|   |      A       |     B      |      C      |
|---|--------------|------------|-------------|
| 1 | `{name}`     | `{surName}`|             |
| 2 | Job: `{job}` |            | `{?job}`    |
| 3 | Sport: `{sport}` |        | `{?sport}`  |
| 4 | Cars of `{name}` `{surName}`: |  |        |
| 5 | > `{brand}`  | `{model}`  | `{#carList}`|
| 6 | &nbsp;&nbsp;`{color}` | `{=year}` | `{/carList}` |
| 7 |              |            |             |
| 8 | Oldest car's year: | `=MIN(B6)` |       |

Then generate a workbook from it:

```ts
import { ExcelTemplater, TemplateData } from 'mitosis-excel-templater-client';

// A local path, an http(s) URL, a Buffer, or a stored { templateId }.
const excelTemplater = new ExcelTemplater('templates/cars.xlsx');

const templateData: TemplateData = {
  name: 'Louis',
  surName: 'Durand',
  job: 'Software Engineer',
  carList: [
    { model: 'C4', color: 'Blue', brand: 'Citroën', year: '2011' },
    { model: '3008', color: 'Brown', brand: 'Peugeot', year: 2012 },
  ],
};

// Writes the file and returns its bytes.
await excelTemplater.saveAsExcel(templateData, 'generated/my-cars.xlsx');
```

`saveAsExcel` always resolves with a `Buffer`, so you can stream it straight to an HTTP response instead of touching the disk:

```ts
const workbook = await excelTemplater.saveAsExcel(templateData);

res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
res.setHeader('Content-Disposition', 'attachment; filename="my-cars.xlsx"');
res.end(workbook);
```

In the result above, row 3 disappears because `sport` is missing, rows 5-6 are repeated for each car, `year` is written as a real number using the cell's number format, and the `=MIN(B6)` formula in row 8 is rewritten to cover every duplicated row.

## Template syntax

| Tag | Meaning |
|---|---|
| `{tagName}` | Replaced by the value of `tagName`. The cell's own styling and format are preserved. |
| `{=tagName}` | Same, but the value is forced to a number so the cell's number format applies. |
| `{?tagName}` | Conditional row: the row is kept only if `tagName` is truthy. The tag itself is removed. |
| `{#tagName}` … `{/tagName}` | Section: everything between the opening and closing row is repeated. |

Sections are flexible:

- an **array** repeats the rows once per element;
- a **single object** renders the rows exactly once;
- an **empty array**, `null` or a missing key removes the rows entirely;
- sections can be **nested** as deeply as you need.

Formatting carried over from the template: cell styles, number and date formats, merged cells, images, named tables, conditional formatting, rich text, hyperlinks and formulas. Row references inside formulas, tables and conditional formatting ranges are **automatically rewritten** as sections expand.

## Reusing stored templates

Uploading the same `.xlsx` on every call is wasteful. Store it once on your account, then reference it by id:

```ts
import { ExcelTemplater, TemplatesClient } from 'mitosis-excel-templater-client';

const templates = new TemplatesClient();

// Once, at deploy time or from a CLI script.
await templates.uploadTemplate('cars', 'templates/cars.xlsx');

// Then, on every request: no upload, smaller and faster calls.
const excelTemplater = new ExcelTemplater({ templateId: 'cars' });
await excelTemplater.saveAsExcel(templateData, 'generated/my-cars.xlsx');
```

The full stored-template API:

```ts
await templates.listTemplates();                      // TemplateSummary[]
await templates.uploadTemplate('cars', buffer);       // upsert by id
await templates.downloadTemplate('cars', 'copy.xlsx');// Buffer, optionally written to disk
await templates.deleteTemplate('cars');
```

## Generating types for your templates

The API can read a template and describe its expected data shape, so your editor catches a typo in a tag name before your users do.

Create a small script, `scripts/gen-types.js`:

```js
const { ExcelTemplater } = require('mitosis-excel-templater-client');

const excelTemplater = new ExcelTemplater('templates/cars.xlsx');
excelTemplater.generateTemplateDataTypescriptFile('src/types/generated/CarsData');
```

Wire it into your build:

```json
{
  "scripts": {
    "gen-types": "node scripts/gen-types.js",
    "build:prod": "npm run gen-types && webpack --mode production"
  }
}
```

And use the generated type:

```ts
import { CarsData } from './types/generated/CarsData';

const templateData: CarsData = {
  name: 'Louis',
  // ...type-checked against the actual template
};
```

Three generators are available:

```ts
// TypeScript type, returned as a string and optionally written to a .d.ts file.
await excelTemplater.generateTemplateDataTypescriptFile('src/types/generated/CarsData');

// JSON Schema (draft 7), for runtime validation with Ajv or for documentation.
await excelTemplater.generateTemplateDataJsonSchema('schemas/cars');

// A ready-to-inspect sample object: strings get their key as value, numbers 0, dates today, booleans true.
const sample = await excelTemplater.generateSampleData();
```

Both file-writing generators accept a second argument, `propsAreOptional`, to mark every property as optional.

Two small helpers are also exposed for cleaning up data coming from an API or a database:

```ts
ExcelTemplater.removeEmptyObjects(data); // recursively drops null/undefined values and empty objects
ExcelTemplater.formatDates(data);        // recursively turns "2023-12-25" strings into Date objects
```

## Configuration

Every constructor accepts an optional second argument:

```ts
const excelTemplater = new ExcelTemplater('templates/cars.xlsx', {
  apiKey: process.env.MY_OWN_VARIABLE,
  timeoutMs: 30_000,
});
```

| Option | Type | Default | Description |
|---|---|---|---|
| `apiKey` | `string` | `MITOSIS_EXCEL_TEMPLATER_API_KEY` env var | Your API key. |
| `baseUrl` | `string` | `https://www.mitosis-excel-templater.dev` | Useful to target a staging environment or a proxy. |
| `timeoutMs` | `number` | `60000` | Per-request timeout. |
| `fetch` | `typeof fetch` | global `fetch` | Custom implementation, for proxies or tests. |

You can also set process-wide defaults once, at startup:

```ts
import { configure } from 'mitosis-excel-templater-client';

configure({ apiKey: process.env.MITOSIS_EXCEL_TEMPLATER_API_KEY, timeoutMs: 30_000 });
```

Resolution order for every option: **constructor argument → `configure()` → environment variable → built-in default**.

## Error handling

All failures throw a subclass of `MitosisApiError`, which exposes the HTTP `status` and the parsed response `body`.

| Error | Status | Raised when |
|---|---|---|
| `AuthenticationError` | 401 | The API key is missing, malformed or unknown. |
| `SubscriptionError` | 403 | Your plan does not cover this resource. |
| `TemplateNotFoundError` | 404 | The `templateId` does not exist on your account. |
| `RateLimitError` | 429 | Your quota for the current period is exhausted. Exposes `remaining`. |
| `TemplateSyntaxError` | — | The template itself is malformed, e.g. a section is opened but never closed. |
| `MitosisApiError` | any | Anything else, including network timeouts (`status: 0`). |

```ts
import { ExcelTemplater, RateLimitError, TemplateSyntaxError } from 'mitosis-excel-templater-client';

try {
  await new ExcelTemplater('templates/cars.xlsx').saveAsExcel(templateData);
} catch (error) {
  if (error instanceof RateLimitError) {
    console.error(`Quota exhausted, ${error.remaining} call(s) left. Upgrade your plan.`);
  } else if (error instanceof TemplateSyntaxError) {
    console.error(`Fix the template: ${error.message}`);
  } else {
    throw error;
  }
}
```

## Rate limits

| Plan | Calls per day |
|---|---|
| Free | 10 |
| Occasional | 200 |
| Unlimited | unlimited |

Current limits and prices are on the [pricing page](https://www.mitosis-excel-templater.dev/pricing). You can check the engine version backing the API at any time, without an API key:

```ts
import { getEngineVersion } from 'mitosis-excel-templater-client';

console.log(await getEngineVersion()); // { name: '@mitosis/mitosis-excel-templater', version: '1.1.1' }
```

## Offline generation, without any API call

If you would rather **not** send your templates and data over the network — or you need to generate files in an air-gapped environment, in a browser without exposing an API key, or without any per-call quota — the underlying engine is also distributed as a private npm package:

**`@mitosis/mitosis-excel-templater`** — the same templating engine, running entirely on your own machine. No API calls, no rate limit, and it works in the browser as well as in Node.js.

It exposes **exactly the same interface** as this client, so switching is a one-line change:

```diff
- import { ExcelTemplater } from 'mitosis-excel-templater-client';
+ import { ExcelTemplater } from '@mitosis/mitosis-excel-templater';
```

To get it:

1. Subscribe to the **npm package** plan on the [pricing page](https://www.mitosis-excel-templater.dev/pricing).
2. Point the `@mitosis` scope at the private registry in your project's `.npmrc`, using your own API key as the auth token:

   ```ini
   @mitosis:registry=https://www.mitosis-excel-templater.dev/registry/
   //www.mitosis-excel-templater.dev/registry/:_authToken=<YOUR_API_KEY>
   ```

3. Install it as usual:

   ```bash
   npm install @mitosis/mitosis-excel-templater
   ```

> ⚠️ Treat that token like any other credential: keep `.npmrc` out of version control, or inject the token from an environment variable in CI.

## Known caveats

- **Charts and graphs are not preserved** in the generated file. Keep them in a separate workbook, or rebuild them from the generated data.
- **Iteration is row-based only** — sections repeat rows, not columns.
- Request size is bounded by the hosted API; very large templates are better stored once with `uploadTemplate` and referenced by id.
- If Excel reports a generated file as needing recovery, the culprit is usually a tag left inside a named table header or a section boundary crossing a merged range.

## Links

- 🌐 **Website** — [mitosis-excel-templater.dev](https://www.mitosis-excel-templater.dev)
- ▶️ **Live demo** — [mitosis-excel-templater.dev](https://www.mitosis-excel-templater.dev)
- 🧪 **API playground** — [mitosis-excel-templater.dev/docs](https://www.mitosis-excel-templater.dev/docs)
- 🔑 **Dashboard & API key** — [mitosis-excel-templater.dev/dashboard](https://www.mitosis-excel-templater.dev/dashboard)
- 💳 **Pricing** — [mitosis-excel-templater.dev/pricing](https://www.mitosis-excel-templater.dev/pricing)
- 🐛 **Issues** — [github.com/Loulou54/mitosis-excel-templater-client/issues](https://github.com/Loulou54/mitosis-excel-templater-client/issues)

## License

[MIT](./LICENSE) © Louis Durand
