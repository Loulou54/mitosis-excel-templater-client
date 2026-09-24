/**
 * A value that can be placed in a single Excel cell.
 *
 * Structurally compatible with ExcelJS' `CellValue`, but declared locally so that this package
 * stays dependency-free: no Excel engine is needed on the client side.
 */
export type CellValue =
  | null
  | undefined
  | string
  | number
  | boolean
  | Date
  | CellRichTextValue
  | CellHyperlinkValue
  | CellFormulaValue
  | CellSharedFormulaValue
  | CellErrorValue;

export type ErrorValue = '#N/A' | '#REF!' | '#NAME?' | '#DIV/0!' | '#NULL!' | '#VALUE!' | '#NUM!';

export interface CellErrorValue {
  error: ErrorValue;
}

export interface RichText {
  text: string;
  font?: Record<string, unknown>;
}

export interface CellRichTextValue {
  richText: RichText[];
}

export interface CellHyperlinkValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

export interface CellFormulaValue {
  formula: string;
  result?: string | number | boolean | Date | CellErrorValue;
  date1904?: boolean;
  shareType?: string;
  ref?: string;
}

export interface CellSharedFormulaValue {
  sharedFormula: string;
  result?: string | number | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

/** A repeatable template section: either a single object, or an array of objects. */
export type Section<T> = T | Array<T>;

/** The structured data fed into an Excel template. */
export type TemplateData = { [key: string]: CellValue | Section<TemplateData> };

/** Reference to a template already stored on your Mitosis Excel Templater account. */
export interface TemplateRef {
  templateId: string;
}

/**
 * Anything accepted as a template: raw bytes, a stored template id, or a string - read from the file
 * system on Node.js, and fetched as a URL in the browser.
 */
export type TemplateSource = Buffer | ArrayBuffer | Uint8Array | string | TemplateRef;

export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

export interface ClientOptions {
  /**
   * Your Mitosis Excel Templater API key.
   * Falls back to `configure({ apiKey })`, then to the `MITOSIS_EXCEL_TEMPLATER_API_KEY` env variable.
   */
  apiKey?: string;
  /** API base URL. Defaults to `https://www.mitosis-excel-templater.dev`. */
  baseUrl?: string;
  /** Custom fetch implementation, mainly useful for tests and proxies. Defaults to the global `fetch`. */
  fetch?: FetchLike;
  /** Per-request timeout in milliseconds. Defaults to 60000. */
  timeoutMs?: number;
}

/** A template stored on your account, as returned by `TemplatesClient.listTemplates()`. */
export interface TemplateSummary {
  id: string;
  templateId: string;
  fileName: string;
  fileSize: number;
  createdAt: string;
}

export interface UploadTemplateResult {
  message: string;
  templateId: string;
  fileName: string;
  fileSize: number;
}

/** Version of the Excel templating engine powering the API. */
export interface EngineVersion {
  name: string;
  version: string;
}
