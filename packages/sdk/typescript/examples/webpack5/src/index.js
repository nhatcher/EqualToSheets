import { initialize } from "@equalto/sheets";

async function run() {
  const { newWorkbook } = await initialize();

  let workbook = newWorkbook("en", "Europe/Warsaw");
  workbook.cell("Sheet1!A1").value = 7;
  workbook.cell("Sheet1!A2").formula = "=A1*2";
  const a2 = workbook.cell("Sheet1!A2").value;

  let div = document.createElement("div");
  div.innerHTML = `A2=${a2}`;
  document.querySelector("body").appendChild(div);
}

run();
