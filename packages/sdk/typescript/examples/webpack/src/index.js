import { JSModel } from "@equalto/wasm";

console.log("a");
let workbook = JSModel.newEmpty("xyz", "en", "Europe/Warsaw");
console.log("b");
workbook.set_input(0, 1, 1, "7");
workbook.set_input(0, 1, 2, "=A1*2");
workbook.evaluate();
const b2 = workbook.get_text_at(0, 1, 2);
let div = document.createElement("div");
div.innerHTML = `B2=${b2}`;
document.querySelector("body").appendChild(div);
console.log("c");
