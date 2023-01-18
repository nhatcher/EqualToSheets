import { Workbook } from "@equalto/wasm";
import { readFileSync } from "fs";

// Note: this doesn't conform to defined TS api yet.
// (Or API will be built on top if it's impossible to define in wasm-bindgen).

let file = readFileSync("./xlsx/test.xlsx");
let model = Workbook.loadFromMemory(file, "en", "Europe/Berlin");

model.evaluate();
console.log(`B1=${model.getFormattedCellValue(0, 1, 2)}`);
console.log(`B2=${model.getFormattedCellValue(0, 2, 2)}`);
console.log(`B3=${model.getFormattedCellValue(0, 3, 2)}`);

console.log("Change B1 to 6");
model.setInput(0, 1, 2, "6");

model.evaluate();
console.log(`B1=${model.getFormattedCellValue(0, 1, 2)}`);
console.log(`B2=${model.getFormattedCellValue(0, 2, 2)}`);
console.log(`B3=${model.getFormattedCellValue(0, 3, 2)}`);

// Exception demo
console.log("");
console.log("---");
console.log("Exception demo");
{
  try {
    // new Workbook("invalid_locale", "invalid_tz"); // Invalid timezone
    Workbook.loadFromMemory(new Uint8Array(0), "en", "Europe/Berlin"); // Invalid Zip Archive
  } catch (e) {
    handleWebAssemblyError(e);
  }
}

function handleWebAssemblyError(e) {
  console.log("error: ", typeof e);
  console.log("instanceof Error", e instanceof Error);
  if (e instanceof Error) {
    console.log("message: ", e.message);
    console.log("parsed message: ", JSON.parse(e.message));
    // next step: construct new error (wrap e for development purposes) and then reraise
  }
}
