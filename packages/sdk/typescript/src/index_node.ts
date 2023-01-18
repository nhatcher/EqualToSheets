import wasm from "./__generated_pkg/equalto_wasm_bg.wasm";
import { setDefaultWasmInit } from "./core";
import { Workbook } from "./core";
import { readFileSync } from "fs";

export { initialize, getApi } from "./core";

declare module "./core" {
  namespace Workbook {
    export function loadFromFileSync(filePath: string): Workbook;
  }
}

Workbook.loadFromFileSync = (filePath) => {
  let file = readFileSync(filePath);
  return Workbook.loadFromMemory(file, "en", "Europe/Berlin");
};

// @ts-ignore
setDefaultWasmInit(() => wasm());
