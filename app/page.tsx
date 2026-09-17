"use client";

import { useMemo, useState } from "react";
import type * as XLSX from "xlsx";
import FileDropzone, { ParsedWorkbook } from "./components/FileDropzone";
import CollectionSummary from "./components/CollectionSummary";
import ManualComparisonTable from "./components/ManualComparisonTable";
import {
  buildExpenseComparisonRows,
  buildManualComparisonRows,
  EXPENSE_METHOD_ORDER,
  ManualDayTotals,
  PAYMENT_METHOD_LABELS,
  PAYMENT_METHOD_ORDER,
  PaymentMethod,
  parseManualStaffReport,
  summarizeCollectionUploads,
  summarizeExpenseReport,
} from "./lib/collectionReport";

const MANUAL_REPORT_CLINICS = [
  { id: "staRosa", label: "DMP Dental Clinic Sta. Rosa" },
  { id: "calamba", label: "DMP Dental Clinic Calamba" },
  { id: "sanPedro", label: "DMP Dental Clinic San Pedro" },
  { id: "stoDomingo", label: "DMP Dental Clinic Sto. Domingo" },
] as const;

type ManualReportClinicId = (typeof MANUAL_REPORT_CLINICS)[number]["id"];

export default function Home() {
  const [paymentUploads, setPaymentUploads] = useState<
    Partial<Record<PaymentMethod, ParsedWorkbook>>
  >({});
  const [expenseUpload, setExpenseUpload] = useState<ParsedWorkbook | null>(
    null
  );
  const [manualReports, setManualReports] = useState<
    Partial<Record<ManualReportClinicId, ParsedWorkbook>>
  >({});

  const collectionDays = useMemo(() => {
    const workbooksByMethod: Partial<Record<PaymentMethod, XLSX.WorkBook>> = {};
    for (const method of PAYMENT_METHOD_ORDER) {
      const upload = paymentUploads[method];
      if (upload) workbooksByMethod[method] = upload.workbook;
    }
    if (Object.keys(workbooksByMethod).length === 0) return null;
    return summarizeCollectionUploads(workbooksByMethod);
  }, [paymentUploads]);

  const expenseReportDays = useMemo(
    () => (expenseUpload ? summarizeExpenseReport(expenseUpload.workbook) : null),
    [expenseUpload]
  );

  const manualReportsByClinicLabel = useMemo(() => {
    const byClinicLabel: Record<string, ManualDayTotals[]> = {};
    for (const clinic of MANUAL_REPORT_CLINICS) {
      const upload = manualReports[clinic.id];
      if (upload) {
        byClinicLabel[clinic.label] = parseManualStaffReport(upload.workbook);
      }
    }
    return byClinicLabel;
  }, [manualReports]);

  const comparisonRows = useMemo(
    () => buildManualComparisonRows(collectionDays ?? [], manualReportsByClinicLabel),
    [collectionDays, manualReportsByClinicLabel]
  );

  const expenseComparisonRows = useMemo(
    () => buildExpenseComparisonRows(expenseReportDays ?? [], manualReportsByClinicLabel),
    [expenseReportDays, manualReportsByClinicLabel]
  );

  const canCompare =
    collectionDays !== null &&
    MANUAL_REPORT_CLINICS.every((clinic) => manualReports[clinic.id]);

  const handlePaymentUpload = (method: PaymentMethod) => (
    parsed: ParsedWorkbook | null
  ) => {
    setPaymentUploads((prev) => {
      const next = { ...prev };
      if (parsed) next[method] = parsed;
      else delete next[method];
      return next;
    });
  };

  const handleExpenseUpload = (parsed: ParsedWorkbook | null) => {
    setExpenseUpload(parsed);
  };

  const handleManualUpload = (clinicId: ManualReportClinicId) => (
    parsed: ParsedWorkbook | null
  ) => {
    setManualReports((prev) => {
      const next = { ...prev };
      if (parsed) next[clinicId] = parsed;
      else delete next[clinicId];
      return next;
    });
  };

  return (
    <div className="flex flex-col flex-1 items-center bg-zinc-50 font-sans dark:bg-black">
      <main className="flex flex-1 w-full max-w-5xl flex-col items-center gap-10 py-32 px-16 bg-white dark:bg-black">
        <div className="flex flex-col items-center gap-2 text-center">
          <h1 className="text-3xl font-semibold leading-10 tracking-tight text-black dark:text-zinc-50">
            Cross-check your reports
          </h1>
          <p className="max-w-md text-sm text-zinc-600 dark:text-zinc-400">
            Upload each payment method report (already filtered) and the
            staff&apos;s manual report to compare them.
          </p>
        </div>

        <div className="grid w-full grid-cols-1 gap-6 sm:grid-cols-2 lg:grid-cols-3">
          {PAYMENT_METHOD_ORDER.map((method) => (
            <FileDropzone
              key={method}
              label={`${PAYMENT_METHOD_LABELS[method]} report`}
              description="Filtered to this payment method"
              onFileParsed={handlePaymentUpload(method)}
            />
          ))}
        </div>
         <div className="grid w-full grid-cols-1 gap-6 sm:grid-cols-2">
          <FileDropzone
            label="Expense report"
            description="From the system, all clinics (Account column determines payment method)"
            onFileParsed={handleExpenseUpload}
          />
        </div>

        <div className="grid w-full grid-cols-1 gap-6 sm:grid-cols-2 lg:grid-cols-4">
          {MANUAL_REPORT_CLINICS.map((clinic) => (
            <FileDropzone
              key={clinic.id}
              label={clinic.label}
              description="Manual report compiled by our staff"
              onFileParsed={handleManualUpload(clinic.id)}
            />
          ))}
        </div>

       

        {/* {collectionDays && (
          <div className="flex w-full flex-col gap-4">
            <h2 className="text-lg font-semibold text-black dark:text-zinc-50">
              Collection Report summary
            </h2>
            <CollectionSummary days={collectionDays} />
          </div>
        )} */}

        {comparisonRows.length > 0 && (
          <div className="flex w-full flex-col gap-4">
            <h2 className="text-lg font-semibold text-black dark:text-zinc-50">
              Manual report comparison (Manual - System)
            </h2>
            <ManualComparisonTable rows={comparisonRows} />
          </div>
        )}

        {expenseComparisonRows.length > 0 && (
          <div className="flex w-full flex-col gap-4">
            <h2 className="text-lg font-semibold text-black dark:text-zinc-50">
              Clinic expenses comparison (Manual - System)
            </h2>
            <ManualComparisonTable rows={expenseComparisonRows} methods={EXPENSE_METHOD_ORDER} />
          </div>
        )}
      </main>
    </div>
  );
}

