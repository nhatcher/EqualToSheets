export enum ErrorKind {
  XlsxError = "XlsxError",
  ReferenceError = "ReferenceError",
  WebAssemblyError = "WebAssemblyError",
  OtherError = "OtherError",
}

export class SheetsError extends Error {
  private _kind: ErrorKind;
  private _wrappedError: Error | null;

  constructor(
    message: string,
    kind: ErrorKind = ErrorKind.OtherError,
    wrappedError?: Error
  ) {
    super(message);
    this.name = "SheetsError";
    this._kind = kind;
    this._wrappedError = wrappedError ?? null;
    Object.setPrototypeOf(this, SheetsError.prototype);
  }

  get kind(): ErrorKind {
    return this._kind;
  }

  get wrappedError(): Error | null {
    return this._wrappedError;
  }
}

/**
 * Wraps a WebAssembly call with try/catch that transforms WASM layer error into `SheetsError`.
 */
export function wrapWebAssemblyCall(wasmCall: () => void) {
  try {
    wasmCall();
  } catch (e) {
    throw wrapWebAssemblyError(e);
  }
}

/**
 * Transforms WASM layer error into `SheetsError`.
 */
export function wrapWebAssemblyError(error: unknown) {
  if (error instanceof SheetsError) {
    throw error;
  }

  if (error instanceof Error) {
    let kind: string;
    let description: string;
    try {
      const parsedError = JSON.parse(error.message);
      if ("kind" in parsedError && "description" in parsedError) {
        kind = parsedError.kind;
        description = parsedError.description;
      } else {
        throw new Error("Could not parse JSON error."); // this error will be immediately consumed
      }
    } catch (parseError) {
      return new SheetsError(error.message, ErrorKind.OtherError, error);
    }

    let errorKind: ErrorKind;
    if (kind === "PlainString") {
      errorKind = ErrorKind.WebAssemblyError;
    } else if (kind === "XlsxError") {
      errorKind = ErrorKind.XlsxError;
    } else {
      errorKind = ErrorKind.OtherError;
    }

    return new SheetsError(description, errorKind, error);
  } else {
    return new SheetsError(`Unexpected error occurred: ${error}`);
  }
}
