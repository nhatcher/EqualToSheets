import wasm from './__generated_pkg/equalto_wasm_bg.wasm';
import { setDefaultWasmInit } from './core';

export { initialize, FormulaErrorCode } from './core';
export type {
  IWorkbook,
  IWorkbookSheets,
  ISheet,
  ICell,
  ICellStyle,
  NavigationDirection,
  FormulaToken,
} from './core';
export { CalcError, ErrorKind } from './errors';

// @ts-ignore
setDefaultWasmInit(() => wasm());
