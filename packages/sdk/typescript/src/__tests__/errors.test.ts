import { ErrorKind, SheetsError, wrapWebAssemblyCall } from "src/errors";

describe("Errors", () => {
  describe("WebAssembly error wrapping", () => {
    test("errors that are already SheetsError are forwarded without change", (done) => {
      const internalError = new Error("internal");
      const thrownError = new SheetsError(
        "",
        ErrorKind.ReferenceError,
        internalError
      );
      try {
        wrapWebAssemblyCall(() => {
          throw thrownError;
        });
        done.fail("Call was expected to throw.");
      } catch (error) {
        expect(error).toBeInstanceOf(SheetsError);
        expect(error).toBe(thrownError);
        expect((error as Error).name).toEqual("SheetsError");
        expect((error as Error).message).toEqual("");
        expect((error as SheetsError).kind).toEqual(ErrorKind.ReferenceError);
        expect((error as SheetsError).wrappedError).toBe(internalError);
        done();
      }
    });

    test("wraps JsError JSON error - PlainString", (done) => {
      try {
        wrapWebAssemblyCall(() => {
          throw Error(
            JSON.stringify({
              kind: "PlainString",
              description: "Something has happened.",
            })
          );
        });
        done.fail("Call was expected to throw.");
      } catch (error) {
        expect(error).toBeInstanceOf(SheetsError);
        expect(error).toBeInstanceOf(Error);
        expect((error as Error).name).toEqual("SheetsError");
        expect((error as Error).message).toEqual("Something has happened.");
        expect((error as SheetsError).kind).toEqual(ErrorKind.WebAssemblyError);
        done();
      }
    });

    test("wraps JsError JSON error - XlsxError", (done) => {
      try {
        wrapWebAssemblyCall(() => {
          throw Error(
            JSON.stringify({
              kind: "XlsxError",
              description: "Something else has happened.",
            })
          );
        });
        done.fail("Call was expected to throw.");
      } catch (error) {
        expect(error).toBeInstanceOf(SheetsError);
        expect(error).toBeInstanceOf(Error);
        expect((error as Error).name).toEqual("SheetsError");
        expect((error as Error).message).toEqual(
          "Something else has happened."
        );
        expect((error as SheetsError).kind).toEqual(ErrorKind.XlsxError);
        done();
      }
    });

    test("wraps JsError JSON error - UnknownError", (done) => {
      try {
        wrapWebAssemblyCall(() => {
          throw Error(
            JSON.stringify({
              kind: "UnknownError",
              description: "Something unknown has happened.",
            })
          );
        });
        done.fail("Call was expected to throw.");
      } catch (error) {
        expect(error).toBeInstanceOf(SheetsError);
        expect(error).toBeInstanceOf(Error);
        expect((error as Error).name).toEqual("SheetsError");
        expect((error as Error).message).toEqual(
          "Something unknown has happened."
        );
        expect((error as SheetsError).kind).toEqual(ErrorKind.OtherError);
        done();
      }
    });

    test("wraps error even if format is unrecognized", (done) => {
      try {
        wrapWebAssemblyCall(() => {
          throw Error("{");
        });
        done.fail("Call was expected to throw.");
      } catch (error) {
        expect(error).toBeInstanceOf(SheetsError);
        expect(error).toBeInstanceOf(Error);
        expect((error as Error).name).toEqual("SheetsError");
        expect((error as Error).message).toEqual("{");
        expect((error as SheetsError).kind).toEqual(ErrorKind.OtherError);
        done();
      }
    });

    test("wraps error even if error type is not Error ", (done) => {
      try {
        wrapWebAssemblyCall(() => {
          throw "I'm a plain string.";
        });
        done.fail("Call was expected to throw.");
      } catch (error) {
        expect(error).toBeInstanceOf(SheetsError);
        expect(error).toBeInstanceOf(Error);
        expect((error as Error).name).toEqual("SheetsError");
        expect((error as Error).message).toEqual(
          "Unexpected error occurred: I'm a plain string."
        );
        expect((error as SheetsError).kind).toEqual(ErrorKind.OtherError);
        done();
      }
    });
  });
});
