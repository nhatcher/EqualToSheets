import init from './__generated_pkg/equalto_wasm';
import { IWorkbook, newWorkbook, loadWorkbookFromMemory } from './api/workbook';
import './dayjsConfig';
import { getFormulaTokens } from './api/utils';

export type { IWorkbook } from './api/workbook';
export type { IWorkbookSheets } from './api/workbookSheets';
export type { ISheet, NavigationDirection } from './api/sheet';
export type { ICell } from './api/cell';
export type { ICellStyle } from './api/style';
export type { FormulaToken } from './api/utils';

export { FormulaErrorCode } from './api/utils';

// Copying type over here directly yields better type generation
export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

let defaultWasmInit: (() => InitInput) | null = null;
export const setDefaultWasmInit = (newDefault: typeof defaultWasmInit) => {
  defaultWasmInit = newDefault;
};

type SheetsApi = {
  newWorkbook(): IWorkbook;
  loadWorkbookFromMemory(data: Uint8Array): IWorkbook;
  utils: {
    getFormulaTokens: typeof getFormulaTokens;
  };
};

async function initializeWasm() {
  // @ts-ignore
  await init(defaultWasmInit());
}

let initializationPromise: Promise<void> | null = null;
export async function initialize(): Promise<SheetsApi> {
  if (initializationPromise === null) {
    initializationPromise = initializeWasm();
  }
  await initializationPromise;
  return {
    newWorkbook,
    loadWorkbookFromMemory,
    utils: {
      getFormulaTokens,
    },
  };
}
