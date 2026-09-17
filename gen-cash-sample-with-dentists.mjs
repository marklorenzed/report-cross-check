import * as XLSX from "xlsx";
import { readFileSync, writeFileSync } from "fs";

const CASH_PATTERN = /\bcash\b/i;
const GCASH_PATTERN = /g\s*-?\s*cash/i;
const isCash = (ref) => {
  const value = String(ref ?? "").trim();
  return CASH_PATTERN.test(value) && !GCASH_PATTERN.test(value);
};

const workbook = XLSX.read(readFileSync("sample-data/report.xlsx"), {
  type: "buffer",
});

const HEADER = ["Date", "Patient", "Clinic/Lab", "Reference #", "Remarks", "Amount"];

function toCashRows(sheetName) {
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  return rows
    .filter((row) => isCash(row["Reference #"]))
    .map((row) => ({
      Date: row["Date"],
      Patient: row["Patient"],
      "Clinic/Lab": row["Clinic/Lab"],
      "Reference #": "Cash",
      Remarks: row["Remarks"],
      Amount: row["Amount"],
    }));
}

const outWorkbook = XLSX.utils.book_new();
const collectionRows = toCashRows("Collection Report");
XLSX.utils.book_append_sheet(
  outWorkbook,
  XLSX.utils.json_to_sheet(collectionRows, { header: HEADER }),
  "Collection Report"
);

const dentistSheetNames = workbook.SheetNames.filter(
  (name) => !["collection report", "collection summary"].includes(name.trim().toLowerCase())
);

let dentistRowTotal = 0;
for (const name of dentistSheetNames) {
  const rows = toCashRows(name);
  dentistRowTotal += rows.reduce((sum, row) => sum + Number(row["Amount"] || 0), 0);
  XLSX.utils.book_append_sheet(
    outWorkbook,
    XLSX.utils.json_to_sheet(rows, { header: [...HEADER, "Procedures"] }),
    name.slice(0, 31)
  );
}

const outBuffer = XLSX.write(outWorkbook, { type: "buffer", bookType: "xlsx" });
writeFileSync("sample-data/cash-report.xlsx", outBuffer);

const collectionTotal = collectionRows.reduce((sum, row) => sum + Number(row["Amount"] || 0), 0);
console.log("Collection Report cash rows:", collectionRows.length, "total:", collectionTotal);
console.log("Dentist sheets:", dentistSheetNames.length, "combined dentist total:", dentistRowTotal);
