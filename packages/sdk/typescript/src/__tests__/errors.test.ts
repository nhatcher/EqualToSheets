import { ErrorKind, SheetsError, wrapWebAssemblyError } from "src/errors";

describe("Errors", () => {
  describe("WebAssembly error wrapping", () => {
    test("errors that are already SheetsError are forwarded without change", () => {
      const internalError = new Error("internal");
      const originalError = new SheetsError(
        "Err!",
        ErrorKind.ReferenceError,
        internalError
      );

      const error = wrapWebAssemblyError(originalError);
      expect(error).toBeInstanceOf(SheetsError);
      expect(error).toBe(originalError);
      expect((error as Error).name).toEqual("SheetsError");
      expect((error as Error).message).toEqual("Err!");
      expect((error as SheetsError).kind).toEqual(ErrorKind.ReferenceError);
      expect((error as SheetsError).wrappedError).toBe(internalError);
    });

    test("wraps JsError JSON error - PlainString", () => {
      const error = wrapWebAssemblyError(
        new Error(
          JSON.stringify({
            kind: "PlainString",
            description: "Something has happened.",
          })
        )
      );
      expect(error).toBeInstanceOf(SheetsError);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).name).toEqual("SheetsError");
      expect((error as Error).message).toEqual("Something has happened.");
      expect((error as SheetsError).kind).toEqual(ErrorKind.WebAssemblyError);
    });

    test("wraps JsError JSON error - XlsxError", () => {
      const error = wrapWebAssemblyError(
        new Error(
          JSON.stringify({
            kind: "XlsxError",
            description: "Something else has happened.",
          })
        )
      );
      expect(error).toBeInstanceOf(SheetsError);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).name).toEqual("SheetsError");
      expect((error as Error).message).toEqual("Something else has happened.");
      expect((error as SheetsError).kind).toEqual(ErrorKind.XlsxError);
    });

    test("wraps JsError JSON error - UnknownError", () => {
      const error = wrapWebAssemblyError(
        new Error(
          JSON.stringify({
            kind: "UnknownError",
            description: "Something unknown has happened.",
          })
        )
      );
      expect(error).toBeInstanceOf(SheetsError);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).name).toEqual("SheetsError");
      expect((error as Error).message).toEqual(
        "Something unknown has happened."
      );
      expect((error as SheetsError).kind).toEqual(ErrorKind.OtherError);
    });

    test("wraps error even if format is unrecognized", () => {
      const error = wrapWebAssemblyError(new Error("{"));
      expect(error).toBeInstanceOf(SheetsError);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).name).toEqual("SheetsError");
      expect((error as Error).message).toEqual("{");
      expect((error as SheetsError).kind).toEqual(ErrorKind.OtherError);
    });

    test("wraps error even if error type is not Error ", () => {
      const error = wrapWebAssemblyError("I'm a plain string.");
      expect(error).toBeInstanceOf(SheetsError);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).name).toEqual("SheetsError");
      expect((error as Error).message).toEqual(
        "Unexpected error occurred: I'm a plain string."
      );
      expect((error as SheetsError).kind).toEqual(ErrorKind.OtherError);
    });
  });
});
