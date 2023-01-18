import { initialize } from "@equalto/sheets";

async function run() {
  const { Workbook } = await initialize();

  let workbook = Workbook.new("en", "Europe/Warsaw");
  workbook.setInput(0, 1, 1, "7");
  workbook.setInput(0, 1, 2, "=A1*2");
  workbook.evaluate();

  const b2 = workbook.getFormattedCellValue(0, 1, 2);

  let div = document.createElement("div");
  div.innerHTML = `B2=${b2}`;
  document.querySelector("body").appendChild(div);
}

run();
