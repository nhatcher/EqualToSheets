import wasm from './__generated_pkg/equalto_wasm_bg.wasm';
import { setDefaultWasmInit } from './core';

export { initialize, getApi } from './core';
export type { IWorkbook, IWorkbookSheets, ISheet, ICell, NavigationDirection } from './core';
export { CalcError, ErrorKind } from './errors';

// @ts-ignore
setDefaultWasmInit(() => wasm());
