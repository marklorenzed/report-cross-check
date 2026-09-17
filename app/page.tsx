"use client";

import { useMemo, useState } from "react";
import type * as XLSX from "xlsx";
import { ParsedWorkbook } from "./components/FileDropzone";
import UploadSidebar from "./components/UploadSidebar";
import Accordion from "./components/Accordion";
import ManualComparisonTable from "./components/ManualComparisonTable";
import { MANUAL_REPORT_CLINICS, ManualReportClinicId } from "./lib/clinics";
import {
  buildExpenseComparisonRows,
  buildManualComparisonRows,
  EXPENSE_METHOD_ORDER,
  ManualDayTotals,
  PAYMENT_METHOD_ORDER,
  PaymentMethod,
  parseManualStaffReport,
  summarizeCollectionUploads,
  summarizeExpenseReport,
} from "./lib/collectionReport";

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
    <div className="flex flex-1 items-stretch bg-zinc-50 font-sans dark:bg-black">
      <UploadSidebar
        onPaymentUpload={handlePaymentUpload}
        onExpenseUpload={handleExpenseUpload}
        onManualUpload={handleManualUpload}
      />

      <main className="flex flex-1 flex-col items-center gap-10 py-16 px-16 bg-white dark:bg-black">
        <div className="flex flex-col items-center gap-2 text-center">
          <h1 className="text-3xl font-semibold leading-10 tracking-tight text-black dark:text-zinc-50">
            Cross-check your reports
          </h1>
          <p className="max-w-md text-sm text-zinc-600 dark:text-zinc-400">
            Upload each payment method report (already filtered) and the
            staff&apos;s manual report to compare them.
          </p>
        </div>

        {comparisonRows.length > 0 && (
          <Accordion title="Manual report comparison (Manual - System)">
            <ManualComparisonTable rows={comparisonRows} />
          </Accordion>
        )}

        {expenseComparisonRows.length > 0 && (
          <Accordion title="Clinic expenses comparison (Manual - System)">
            <ManualComparisonTable rows={expenseComparisonRows} methods={EXPENSE_METHOD_ORDER} />
          </Accordion>
        )}
      </main>
    </div>
  );
}

