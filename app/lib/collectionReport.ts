import * as XLSX from "xlsx";
import { findSheetByName } from "./xlsx";

export type PaymentMethod =
  | "cash"
  | "gcash"
  | "maya"
  | "bankTransfer"
  | "cc"
  | "dc";

export const PAYMENT_METHOD_ORDER: PaymentMethod[] = [
  "cash",
  "gcash",
  "maya",
  "bankTransfer",
  "cc",
  "dc",
];

export const PAYMENT_METHOD_LABELS: Record<PaymentMethod, string> = {
  cash: "Cash",
  gcash: "GCash",
  maya: "Maya",
  bankTransfer: "Bank Transfer",
  cc: "Credit Card",
  dc: "Debit Card",
};

export type CollectionRow = {
  date: string;
  clinicLab: string;
  amount: number;
};

type CollectionRawRow = {
  Date?: unknown;
  "Clinic/Lab"?: unknown;
  Amount?: unknown;
};

function parseSheetRows(sheet: XLSX.WorkSheet): CollectionRow[] {
  const rawRows = XLSX.utils.sheet_to_json<CollectionRawRow>(sheet);

  return rawRows
    .map((row) => {
      const date = String(row["Date"] ?? "").trim();
      const clinicLab = String(row["Clinic/Lab"] ?? "").trim();
      const amount = Number(row["Amount"]);
      return {
        date,
        clinicLab,
        amount: Number.isFinite(amount) ? amount : 0,
      };
    })
    .filter((row) => row.date && row.clinicLab);
}

/** Reads a payment-method-filtered export into normalized rows. Falls back to the first sheet if there's no "Collection Report" sheet. */
export function parsePaymentMethodSheet(
  workbook: XLSX.WorkBook
): CollectionRow[] {
  const sheet =
    findSheetByName(workbook, "Collection Report") ??
    workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];

  return parseSheetRows(sheet);
}

const NON_DENTIST_SHEET_NAMES = new Set(["collection report", "collection summary"]);

export type PaymentMethodTotals = Record<PaymentMethod, number>;

function emptyTotals(): PaymentMethodTotals {
  return { cash: 0, gcash: 0, maya: 0, bankTransfer: 0, cc: 0, dc: 0 };
}

/** Every sheet other than "Collection Report"/"Collection Summary" is named after a dentist and lists only their rows. */
export function parseDentistTotalsByDate(
  workbook: XLSX.WorkBook,
  method: PaymentMethod
): Map<string, Map<string, PaymentMethodTotals>> {
  const totalsByDate = new Map<string, Map<string, PaymentMethodTotals>>();

  for (const sheetName of workbook.SheetNames) {
    if (NON_DENTIST_SHEET_NAMES.has(sheetName.trim().toLowerCase())) continue;

    const rows = parseSheetRows(workbook.Sheets[sheetName]);
    for (const row of rows) {
      let dentistMap = totalsByDate.get(row.date);
      if (!dentistMap) {
        dentistMap = new Map();
        totalsByDate.set(row.date, dentistMap);
      }
      let totals = dentistMap.get(sheetName);
      if (!totals) {
        totals = emptyTotals();
        dentistMap.set(sheetName, totals);
      }
      totals[method] += row.amount;
    }
  }

  return totalsByDate;
}

function mergeDentistTotalsByDate(
  maps: Map<string, Map<string, PaymentMethodTotals>>[]
): Map<string, Map<string, PaymentMethodTotals>> {
  const merged = new Map<string, Map<string, PaymentMethodTotals>>();

  for (const map of maps) {
    for (const [date, dentistMap] of map) {
      let targetDentistMap = merged.get(date);
      if (!targetDentistMap) {
        targetDentistMap = new Map();
        merged.set(date, targetDentistMap);
      }
      for (const [dentist, totals] of dentistMap) {
        let targetTotals = targetDentistMap.get(dentist);
        if (!targetTotals) {
          targetTotals = emptyTotals();
          targetDentistMap.set(dentist, targetTotals);
        }
        for (const method of PAYMENT_METHOD_ORDER) {
          targetTotals[method] += totals[method];
        }
      }
    }
  }

  return merged;
}

export type ClinicSummary = {
  clinicLab: string;
  totals: PaymentMethodTotals;
  total: number;
};

export type DentistTotal = {
  name: string;
  totals: PaymentMethodTotals;
  total: number;
};

export type DaySummary = {
  date: string;
  clinics: ClinicSummary[];
  dentists: DentistTotal[];
  totals: PaymentMethodTotals;
  total: number;
};

/** Merges per-payment-method uploads (each already filtered to one method) into a day/clinic/dentist summary. */
export function summarizeCollectionUploads(
  uploads: Partial<Record<PaymentMethod, XLSX.WorkBook>>
): DaySummary[] {
  const dayMap = new Map<string, Map<string, ClinicSummary>>();
  const dentistTotalMaps: Map<string, Map<string, PaymentMethodTotals>>[] = [];

  for (const method of PAYMENT_METHOD_ORDER) {
    const workbook = uploads[method];
    if (!workbook) continue;

    const rows = parsePaymentMethodSheet(workbook);
    for (const row of rows) {
      let clinicMap = dayMap.get(row.date);
      if (!clinicMap) {
        clinicMap = new Map();
        dayMap.set(row.date, clinicMap);
      }

      let clinic = clinicMap.get(row.clinicLab);
      if (!clinic) {
        clinic = { clinicLab: row.clinicLab, totals: emptyTotals(), total: 0 };
        clinicMap.set(row.clinicLab, clinic);
      }

      clinic.totals[method] += row.amount;
      clinic.total += row.amount;
    }

    dentistTotalMaps.push(parseDentistTotalsByDate(workbook, method));
  }

  const dentistTotalsByDate = mergeDentistTotalsByDate(dentistTotalMaps);
  const dates = new Set([...dayMap.keys(), ...dentistTotalsByDate.keys()]);

  const days: DaySummary[] = [];
  for (const date of dates) {
    const clinicMap = dayMap.get(date);
    const clinics = clinicMap
      ? [...clinicMap.values()].sort((a, b) =>
          a.clinicLab.localeCompare(b.clinicLab)
        )
      : [];

    const totals = emptyTotals();
    let total = 0;
    for (const clinic of clinics) {
      for (const method of PAYMENT_METHOD_ORDER) {
        totals[method] += clinic.totals[method];
      }
      total += clinic.total;
    }

    const dentistTotals = dentistTotalsByDate.get(date);
    const dentists = dentistTotals
      ? [...dentistTotals.entries()]
          .map(([name, totals]) => ({
            name,
            totals,
            total: PAYMENT_METHOD_ORDER.reduce(
              (sum, method) => sum + totals[method],
              0
            ),
          }))
          .sort((a, b) => b.total - a.total)
      : [];

    days.push({ date, clinics, dentists, totals, total });
  }

  return days.sort((a, b) => parseReportDate(a.date) - parseReportDate(b.date));
}

function parseReportDate(value: string): number {
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? 0 : parsed;
}
