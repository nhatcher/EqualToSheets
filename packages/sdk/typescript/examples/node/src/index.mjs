import { JSModel } from "@equalto/wasm";
import { readFileSync } from "fs";

// Note: this doesn't conform to defined TS api yet.
// (Or API will be built on top if it's impossible to define in wasm-bindgen).

let file = readFileSync("./xlsx/test.xlsx");
let model = JSModel.newFromExcelFile("example", file, "en", "UTC");

model.evaluate();
console.log(`B1=${model.get_text_at(0, 1, 2)}`);
console.log(`B2=${model.get_text_at(0, 2, 2)}`);
console.log(`B3=${model.get_text_at(0, 3, 2)}`);

console.log("Change B1 to 6");
model.set_input(0, 1, 2, "6");

model.evaluate();
console.log(`B1=${model.get_text_at(0, 1, 2)}`);
console.log(`B2=${model.get_text_at(0, 2, 2)}`);
console.log(`B3=${model.get_text_at(0, 3, 2)}`);
