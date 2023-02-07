export enum ErrorKind {
  XlsxError = 'XlsxError',
  ReferenceError = 'ReferenceError',
  WebAssemblyError = 'WebAssemblyError',
  TypeError = 'TypeError',
  OtherError = 'OtherError',
}

export class CalcError extends Error {
  private _kind: ErrorKind;
  private _wrappedError: Error | null;

  constructor(message: string, kind: ErrorKind = ErrorKind.OtherError, wrappedError?: Error) {
    super(message);
    this.name = 'CalcError';
    this._kind = kind;
    this._wrappedError = wrappedError ?? null;
    Object.setPrototypeOf(this, CalcError.prototype);
  }

  get kind(): ErrorKind {
    return this._kind;
  }

  get wrappedError(): Error | null {
    return this._wrappedError;
  }
}

/**
 * Transforms WASM layer error into `CalcError`.
 */
export function wrapWebAssemblyError(error: unknown) {
  if (error instanceof CalcError) {
    return error;
  }

  if (error instanceof Error) {
    let kind: string;
    let description: string;
    try {
      const parsedError = JSON.parse(error.message);
      if ('kind' in parsedError && 'description' in parsedError) {
        kind = parsedError.kind;
        description = parsedError.description;
      } else {
        throw new Error('Could not parse JSON error.'); // this error will be immediately consumed
      }
    } catch (parseError) {
      return new CalcError(error.message, ErrorKind.OtherError, error);
    }

    let errorKind: ErrorKind;
    if (kind === 'PlainString') {
      errorKind = ErrorKind.WebAssemblyError;
    } else if (kind === 'XlsxError') {
      errorKind = ErrorKind.XlsxError;
    } else {
      errorKind = ErrorKind.OtherError;
    }

    return new CalcError(description, errorKind, error);
  } else {
    return new CalcError(`Unexpected error occurred: ${error}`);
  }
}
