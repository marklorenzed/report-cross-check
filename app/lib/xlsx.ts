import * as XLSX from "xlsx";

/** Finds a sheet by name, case-insensitively, falling back to a partial match. */
export function findSheetByName(
  workbook: XLSX.WorkBook,
  name: string
): XLSX.WorkSheet | null {
  const target = name.trim().toLowerCase();
  const exact = workbook.SheetNames.find(
    (sheetName) => sheetName.trim().toLowerCase() === target
  );
  const partial = workbook.SheetNames.find((sheetName) =>
    sheetName.trim().toLowerCase().includes(target)
  );
  const match = exact ?? partial;
  return match ? workbook.Sheets[match] : null;
}
