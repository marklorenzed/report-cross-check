import * as XLSX from "xlsx";
import { readFileSync } from "fs";

const workbook = XLSX.read(readFileSync("sample-data/report.xlsx"), {
  type: "buffer",
});

function dump(sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    console.log(`--- ${sheetName}: NOT FOUND ---`);
    return;
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  console.log(`--- ${sheetName} ---`);
  console.log(JSON.stringify(rows.slice(0, 6), null, 2));
}

dump("Dona Marie Dumalaon");
dump("Collection Report");
dump("Collection Summary");
