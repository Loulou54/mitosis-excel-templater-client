/** Base class for every error raised by the Mitosis Excel Templater API. */
export class MitosisApiError extends Error {
  /** HTTP status returned by the API, or 0 when the request never completed. */
  public readonly status: number;
  /** Raw parsed response body, when the API returned one. */
  public readonly body: unknown;

  constructor(message: string, status: number, body?: unknown) {
    super(message);
    this.name = new.target.name;
    this.status = status;
    this.body = body;
  }
}

/** The API key is missing, malformed or unknown (HTTP 401). */
export class AuthenticationError extends MitosisApiError {
  constructor(message: string, body?: unknown) {
    super(message, 401, body);
  }
}

/** The account lacks the subscription required for this resource (HTTP 403). */
export class SubscriptionError extends MitosisApiError {
  constructor(message: string, body?: unknown) {
    super(message, 403, body);
  }
}

/** The requested stored template does not exist (HTTP 404). */
export class TemplateNotFoundError extends MitosisApiError {
  constructor(message: string, body?: unknown) {
    super(message, 404, body);
  }
}

/** The plan quota for the current period is exhausted (HTTP 429). */
export class RateLimitError extends MitosisApiError {
  /** Calls left for the current period, as reported by the API. */
  public readonly remaining: number;

  constructor(message: string, body?: unknown, remaining = 0) {
    super(message, 429, body);
    this.remaining = remaining;
  }
}

/** The Excel template itself is malformed, e.g. a section is opened but never closed. */
export class TemplateSyntaxError extends MitosisApiError {}

function extractMessage(body: unknown): string | undefined {
  if (typeof body === 'string' && body.trim()) {
    return body.trim();
  }
  if (body && typeof body === 'object' && 'error' in body) {
    const error = (body as { error: unknown }).error;
    if (typeof error === 'string' && error.trim()) {
      return error.trim();
    }
  }
  return undefined;
}

function extractRemaining(body: unknown): number {
  if (body && typeof body === 'object' && 'remaining' in body) {
    const remaining = (body as { remaining: unknown }).remaining;
    if (typeof remaining === 'number') {
      return remaining;
    }
  }
  return 0;
}

/** Turns an API error response into the most specific error class available. */
export function fromResponse(status: number, body: unknown): MitosisApiError {
  const message = extractMessage(body) ?? `Request failed with status ${status}`;

  if (/template formatting error/i.test(message)) {
    return new TemplateSyntaxError(message, status, body);
  }

  switch (status) {
    case 401:
      return new AuthenticationError(message, body);
    case 403:
      return new SubscriptionError(message, body);
    case 404:
      return new TemplateNotFoundError(message, body);
    case 429:
      return new RateLimitError(message, body, extractRemaining(body));
    default:
      return new MitosisApiError(message, status, body);
  }
}
