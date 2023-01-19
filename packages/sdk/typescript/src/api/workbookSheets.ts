import { WasmWorkbook } from "../__generated_pkg/equalto_wasm";
import { Sheet, ISheet } from "./sheet";

export interface IWorkbookSheets {
  add(sheetName?: string): ISheet;
  get(sheetName: string): ISheet;
  get(sheetIndex: number): ISheet;
  get(sheetRef: string | number): ISheet;
  all(): ISheet[];
}

export class WorkbookSheets implements IWorkbookSheets {
  private readonly _wasmWorkbook: WasmWorkbook;
  private _sheetLookups: SheetLookup;

  constructor(wasmWorkbook: WasmWorkbook) {
    this._wasmWorkbook = wasmWorkbook;
    this._sheetLookups = loadSheetLookups(this._wasmWorkbook);
  }

  add(sheetName?: string): ISheet {
    if (sheetName !== undefined) {
      this._wasmWorkbook.addSheet(sheetName);
    } else {
      this._wasmWorkbook.newSheet();
    }
    this._refreshSheetLookups();
    return this.get(this._sheetLookups.numSheets - 1);
  }

  get(sheetName: string): ISheet;
  get(sheetIndex: number): ISheet;
  get(sheetRef: string | number): ISheet {
    let sheetId: number;
    if (typeof sheetRef === "string") {
      sheetId = this._getSheetIdBySheetName(sheetRef);
    } else if (typeof sheetRef === "number") {
      sheetId = this._getSheetIdBySheetIndex(sheetRef);
    } else {
      throw new Error("Sheet reference must be either string or number.");
    }

    return new Sheet(this, this._wasmWorkbook, sheetId);
  }

  all(): ISheet[] {
    const sheets = [];
    for (let index = 0; index < this._sheetLookups.numSheets; index += 1) {
      sheets.push(this.get(index));
    }
    return sheets;
  }

  _refreshSheetLookups() {
    this._sheetLookups = loadSheetLookups(this._wasmWorkbook);
  }

  _getSheetIndexBySheetId(sheetId: number): number {
    if (sheetId in this._sheetLookups.sheetIdToSheetIndex) {
      return this._sheetLookups.sheetIdToSheetIndex[sheetId];
    }
    throw new Error(`Could not find sheet with sheetId=${sheetId}`);
  }

  _getSheetNameBySheetId(sheetId: number): string {
    if (sheetId in this._sheetLookups.sheetIdToSheetName) {
      return this._sheetLookups.sheetIdToSheetName[sheetId];
    }
    throw new Error(`Could not find sheet with sheetId=${sheetId}`);
  }

  private _getSheetIdBySheetName(sheetName: string): number {
    if (sheetName in this._sheetLookups.sheetNameToSheetId) {
      return this._sheetLookups.sheetNameToSheetId[sheetName];
    }
    throw new Error(`Could not find sheet with name="${sheetName}"`);
  }

  private _getSheetIdBySheetIndex(sheetIndex: number): number {
    if (sheetIndex in this._sheetLookups.sheetIndexToSheetId) {
      return this._sheetLookups.sheetIndexToSheetId[sheetIndex];
    }
    throw new Error(`Could not find sheet at index=${sheetIndex}`);
  }
}

type SheetLookup = {
  numSheets: number;
  sheetIndexToSheetId: Record<number, number>;
  sheetNameToSheetId: Record<string, number>;
  sheetIdToSheetIndex: Record<number, number>;
  sheetIdToSheetName: Record<number, string>;
};

function loadSheetLookups(wasmWorkbook: WasmWorkbook): SheetLookup {
  const worksheetIds = wasmWorkbook.getWorksheetIds() as number[];
  const worksheetNames = wasmWorkbook.getWorksheetNames() as string[];

  if (worksheetIds.length !== worksheetNames.length) {
    throw new Error(
      "Internal error. Number of worksheet names does not match number of worksheet IDs"
    );
  }

  const sheetIndexToSheetId: Record<number, number> = {};
  const sheetNameToSheetId: Record<string, number> = {};
  const sheetIdToSheetIndex: Record<number, number> = {};
  const sheetIdToSheetName: Record<number, string> = {};

  let length = worksheetIds.length;
  for (let index = 0; index < length; index += 1) {
    let id = worksheetIds[index];
    let name = worksheetNames[index];
    sheetIndexToSheetId[index] = id;
    sheetNameToSheetId[name] = id;
    sheetIdToSheetIndex[id] = index;
    sheetIdToSheetName[id] = name;
  }

  return {
    numSheets: length,
    sheetIndexToSheetId,
    sheetNameToSheetId,
    sheetIdToSheetIndex,
    sheetIdToSheetName,
  };
}
