import init from "./__generated_pkg/equalto_wasm";
import { IWorkbook, newWorkbook, loadWorkbookFromMemory } from "./api/workbook";
import "./dayjsConfig";

export type { IWorkbook } from "./api/workbook";
export type { IWorkbookSheets } from "./api/workbookSheets";
export type { ISheet } from "./api/sheet";
export type { ICell } from "./api/cell";

// Copying type over here directly yields better type generation
export type InitInput =
  | RequestInfo
  | URL
  | Response
  | BufferSource
  | WebAssembly.Module;

let defaultWasmInit: (() => InitInput) | null = null;
export const setDefaultWasmInit = (newDefault: typeof defaultWasmInit) => {
  defaultWasmInit = newDefault;
};

type SheetsApi = {
  newWorkbook(locale: string, timezone: string): IWorkbook;
  loadWorkbookFromMemory(
    data: Uint8Array,
    locale: string,
    timezone: string
  ): IWorkbook;
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
    newWorkbook,
    loadWorkbookFromMemory,
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
    newWorkbook,
    loadWorkbookFromMemory,
  };
}
