import wasm from './__generated_pkg/equalto_wasm_bg.wasm';
import { setDefaultWasmInit } from './core';

export { initialize } from './core';
export type {
  IWorkbook,
  IWorkbookSheets,
  ISheet,
  ICell,
  ICellStyle,
  CellStyleSnapshot,
  CellStyleUpdateValues,
  NavigationDirection,
  FormulaToken,
  FormulaErrorCode,
} from './core';
export { CalcError, ErrorKind } from './errors';
export { convertSpreadsheetDateToISOString } from './utils';

// @ts-ignore
setDefaultWasmInit(() => wasm());
