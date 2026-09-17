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

const MANUAL_REPORT_HEADER_ALIASES: Record<PaymentMethod, string[]> = {
  cash: ["cash"],
  gcash: ["gcash"],
  maya: ["maya", "paymaya"],
  bankTransfer: ["bank", "bank transfer", "bdo", "bpi"],
  cc: ["cc", "credit card"],
  dc: ["dc", "debit card"],
};

function excelSerialToDateLabel(serial: number): string {
  const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
    timeZone: "UTC",
  });
}

function findManualReportHeaderRowIndex(rows: unknown[][]): number {
  const paymentLabels = new Set(
    Object.values(MANUAL_REPORT_HEADER_ALIASES).flat()
  );
  return rows.findIndex((row) =>
    row.some((cell) => {
      const label = String(cell ?? "").trim().toLowerCase();
      return label === "dentist" || paymentLabels.has(label);
    })
  );
}

function mapManualReportColumns(
  headerRow: unknown[]
): Record<PaymentMethod, number[]> {
  const columns: Record<PaymentMethod, number[]> = {
    cash: [],
    gcash: [],
    maya: [],
    bankTransfer: [],
    cc: [],
    dc: [],
  };
  headerRow.forEach((cell, index) => {
    const label = String(cell ?? "").trim().toLowerCase();
    if (!label) return;
    for (const method of PAYMENT_METHOD_ORDER) {
      if (MANUAL_REPORT_HEADER_ALIASES[method].includes(label)) {
        columns[method].push(index);
      }
    }
  });
  return columns;
}

export type ManualDayTotals = {
  date: string;
  totals: PaymentMethodTotals;
  expenses: PaymentMethodTotals;
};

function readRowTotals(
  row: unknown[],
  columns: Record<PaymentMethod, number[]>
): PaymentMethodTotals {
  const totals = emptyTotals();
  for (const method of PAYMENT_METHOD_ORDER) {
    totals[method] = columns[method].reduce((sum, columnIndex) => {
      const value = Number(row[columnIndex]);
      return sum + (Number.isFinite(value) ? value : 0);
    }, 0);
  }
  return totals;
}

/**
 * Manual staff reports list one dentist per row per day (day starts with a
 * date serial in column A), ending in a TOTAL row and an EXPENSES row. Only
 * those two rows are used; the per-dentist rows and END TOTAL row are
 * ignored, as is column A on the TOTAL/EXPENSES rows (it holds a grand
 * total, not a date).
 */
export function parseManualStaffReport(
  workbook: XLSX.WorkBook
): ManualDayTotals[] {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1 });
  const headerIndex = findManualReportHeaderRowIndex(rows);
  if (headerIndex === -1) return [];

  const columns = mapManualReportColumns(rows[headerIndex]);
  const days: ManualDayTotals[] = [];
  let pendingDateSerial: number | null = null;
  let pendingTotals: PaymentMethodTotals | null = null;

  const flushPendingDay = () => {
    if (pendingDateSerial !== null && pendingTotals !== null) {
      days.push({
        date: excelSerialToDateLabel(pendingDateSerial),
        totals: pendingTotals,
        expenses: emptyTotals(),
      });
    }
    pendingDateSerial = null;
    pendingTotals = null;
  };

  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const label = String(row[1] ?? "").trim().toUpperCase();

    if (label === "TOTAL") {
      if (pendingDateSerial !== null) pendingTotals = readRowTotals(row, columns);
      continue;
    }

    if (label === "EXPENSES") {
      if (pendingDateSerial !== null && pendingTotals !== null) {
        days.push({
          date: excelSerialToDateLabel(pendingDateSerial),
          totals: pendingTotals,
          expenses: readRowTotals(row, columns),
        });
        pendingDateSerial = null;
        pendingTotals = null;
      }
      continue;
    }

    if (label === "END TOTAL") continue;

    if (typeof row[0] === "number") {
      flushPendingDay();
      pendingDateSerial = row[0];
    }
  }

  flushPendingDay();

  return days;
}

export type ManualComparisonRow = {
  date: string;
  clinicLabel: string;
  reportTotals: PaymentMethodTotals | null;
  manualTotals: PaymentMethodTotals | null;
};

const COMPARISON_KEY_SEPARATOR = "\u0000";

type DatedTotals = { date: string; totals: PaymentMethodTotals };

/** Joins a payment-method-report summary with per-clinic manual totals by date + clinic. */
function joinReportWithManualTotals(
  days: DaySummary[],
  manualTotalsByClinicLabel: Record<string, DatedTotals[]>
): ManualComparisonRow[] {
  const reportTotals = new Map<string, PaymentMethodTotals>();
  for (const day of days) {
    for (const clinic of day.clinics) {
      reportTotals.set(
        `${day.date}${COMPARISON_KEY_SEPARATOR}${clinic.clinicLab}`,
        clinic.totals
      );
    }
  }

  const manualTotals = new Map<string, PaymentMethodTotals>();
  for (const [clinicLabel, entries] of Object.entries(
    manualTotalsByClinicLabel
  )) {
    for (const entry of entries) {
      manualTotals.set(
        `${entry.date}${COMPARISON_KEY_SEPARATOR}${clinicLabel}`,
        entry.totals
      );
    }
  }

  const keys = new Set([...reportTotals.keys(), ...manualTotals.keys()]);
  const rows: ManualComparisonRow[] = [...keys].map((key) => {
    const [date, clinicLabel] = key.split(COMPARISON_KEY_SEPARATOR);
    return {
      date,
      clinicLabel,
      reportTotals: reportTotals.get(key) ?? null,
      manualTotals: manualTotals.get(key) ?? null,
    };
  });

  return rows.sort((a, b) => {
    const dateDiff = parseReportDate(a.date) - parseReportDate(b.date);
    if (dateDiff !== 0) return dateDiff;
    return a.clinicLabel.localeCompare(b.clinicLabel);
  });
}

/** Joins the payment-method-report summary with per-clinic manual reports by date + clinic. */
export function buildManualComparisonRows(
  days: DaySummary[],
  manualReportsByClinicLabel: Record<string, ManualDayTotals[]>
): ManualComparisonRow[] {
  const manualTotalsByClinicLabel: Record<string, DatedTotals[]> = {};
  for (const [clinicLabel, manualDays] of Object.entries(
    manualReportsByClinicLabel
  )) {
    manualTotalsByClinicLabel[clinicLabel] = manualDays.map((day) => ({
      date: day.date,
      totals: day.totals,
    }));
  }
  return joinReportWithManualTotals(days, manualTotalsByClinicLabel);
}

export const EXPENSE_METHOD_ORDER: PaymentMethod[] = [
  "cash",
  "gcash",
  "bankTransfer",
  "maya",
];

/** Joins a system expense-report summary with per-clinic manual EXPENSES rows by date + clinic. */
export function buildExpenseComparisonRows(
  expenseReportDays: DaySummary[],
  manualReportsByClinicLabel: Record<string, ManualDayTotals[]>
): ManualComparisonRow[] {
  const manualExpensesByClinicLabel: Record<string, DatedTotals[]> = {};
  for (const [clinicLabel, manualDays] of Object.entries(
    manualReportsByClinicLabel
  )) {
    manualExpensesByClinicLabel[clinicLabel] = manualDays.map((day) => ({
      date: day.date,
      totals: day.expenses,
    }));
  }
  return joinReportWithManualTotals(expenseReportDays, manualExpensesByClinicLabel);
}

const EXPENSE_ACCOUNT_METHODS: Record<string, PaymentMethod> = {
  "dmp cash": "cash",
  cash: "cash",
  "dmp gcash": "gcash",
  gcash: "gcash",
  "dmp maya": "maya",
  "dmp paymaya": "maya",
  maya: "maya",
  paymaya: "maya",
  "dmp bank account": "bankTransfer",
  "dmp bank": "bankTransfer",
  "bank account": "bankTransfer",
  "bank transfer": "bankTransfer",
  bank: "bankTransfer",
  "dmp cc": "cc",
  "credit card": "cc",
  "dmp dc": "dc",
  "debit card": "dc",
};

type ExpenseReportRawRow = {
  Date?: unknown;
  Clinic?: unknown;
  Account?: unknown;
  Amount?: unknown;
};

/**
 * The system expense report lists every clinic's expenses in one sheet; the
 * "Account" column (e.g. "DMP CASH", "DMP GCASH", "DMP BANK ACCOUNT")
 * determines which payment method each row's amount counts against.
 */
export function summarizeExpenseReport(workbook: XLSX.WorkBook): DaySummary[] {
  const sheet =
    findSheetByName(workbook, "Expense Report") ??
    workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];

  const rawRows = XLSX.utils.sheet_to_json<ExpenseReportRawRow>(sheet);
  const dayMap = new Map<string, Map<string, ClinicSummary>>();

  for (const row of rawRows) {
    const date = String(row["Date"] ?? "").trim();
    const clinicLab = String(row["Clinic"] ?? "").trim();
    const account = String(row["Account"] ?? "").trim().toLowerCase();
    const method = EXPENSE_ACCOUNT_METHODS[account];
    const amount = Number(row["Amount"]);
    if (!date || !clinicLab || !method) continue;

    let clinicMap = dayMap.get(date);
    if (!clinicMap) {
      clinicMap = new Map();
      dayMap.set(date, clinicMap);
    }

    let clinic = clinicMap.get(clinicLab);
    if (!clinic) {
      clinic = { clinicLab, totals: emptyTotals(), total: 0 };
      clinicMap.set(clinicLab, clinic);
    }

    const value = Number.isFinite(amount) ? amount : 0;
    clinic.totals[method] += value;
    clinic.total += value;
  }

  const days: DaySummary[] = [];
  for (const [date, clinicMap] of dayMap) {
    const clinics = [...clinicMap.values()].sort((a, b) =>
      a.clinicLab.localeCompare(b.clinicLab)
    );

    const totals = emptyTotals();
    let total = 0;
    for (const clinic of clinics) {
      for (const method of PAYMENT_METHOD_ORDER) totals[method] += clinic.totals[method];
      total += clinic.total;
    }

    days.push({ date, clinics, dentists: [], totals, total });
  }

  return days.sort((a, b) => parseReportDate(a.date) - parseReportDate(b.date));
}

