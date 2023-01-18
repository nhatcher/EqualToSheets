import init, { InitInput, WasmWorkbook } from "./__generated_pkg/equalto_wasm";

let defaultWasmInit: (() => InitInput) | null = null;
export const setDefaultWasmInit = (newDefault: typeof defaultWasmInit) => {
  defaultWasmInit = newDefault;
};

export class Workbook {
  private _wasmWorkbook: WasmWorkbook;

  private constructor(wasmWorkbook: WasmWorkbook) {
    this._wasmWorkbook = wasmWorkbook;
  }

  static new(locale: string, timezone: string): Workbook {
    return new Workbook(new WasmWorkbook(locale, timezone));
  }

  static loadFromMemory(
    data: Uint8Array,
    locale: string,
    timezone: string
  ): Workbook {
    return new Workbook(WasmWorkbook.loadFromMemory(data, locale, timezone));
  }

  evaluate() {
    this._wasmWorkbook.evaluate();
  }

  setInput(
    sheet: number,
    row: number,
    column: number,
    value: string,
    style: number
  ) {
    this._wasmWorkbook.setInput(sheet, row, column, value, style);
  }

  getFormattedCellValue(sheet: number, row: number, column: number): string {
    return this._wasmWorkbook.getFormattedCellValue(sheet, row, column);
  }
}

type SheetsApi = {
  Workbook: typeof Workbook;
};

async function initializeWasm() {
  // @ts-ignore
  await init(defaultWasmInit());
}

let initializationPromise: Promise<void> | null = null;
export async function initialize(): Promise<SheetsApi> {
  if (initializationPromise !== null) {
    throw new Error("Sheets API cannot be initialized twice.");
  }
  initializationPromise = initializeWasm();
  await initializationPromise;
  return {
    Workbook,
  };
}

/**
 * Requires previous module initialization using `initialize`. Initialization
 * does not have to complete, but must be started.
 */
export async function getApi(): Promise<SheetsApi> {
  if (initializationPromise === null) {
    throw new Error(
      "Sheets API needs to be initialized prior to requesting API."
    );
  }
  await initializationPromise;
  return {
    Workbook,
  };
}
