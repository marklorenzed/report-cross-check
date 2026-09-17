"use client";

import { useState } from "react";
import FileDropzone, { ParsedWorkbook } from "./FileDropzone";
import { MANUAL_REPORT_CLINICS, ManualReportClinicId } from "../lib/clinics";
import {
  PAYMENT_METHOD_LABELS,
  PAYMENT_METHOD_ORDER,
  PaymentMethod,
} from "../lib/collectionReport";

type UploadSidebarProps = {
  onPaymentUpload: (method: PaymentMethod) => (parsed: ParsedWorkbook | null) => void;
  onExpenseUpload: (parsed: ParsedWorkbook | null) => void;
  onManualUpload: (clinicId: ManualReportClinicId) => (parsed: ParsedWorkbook | null) => void;
};

export default function UploadSidebar({
  onPaymentUpload,
  onExpenseUpload,
  onManualUpload,
}: UploadSidebarProps) {
  const [collapsed, setCollapsed] = useState(false);

  if (collapsed) {
    return (
      <aside className="sticky top-0 flex h-screen w-12 shrink-0 flex-col items-center border-r border-black/[.08] bg-white py-4 dark:border-white/[.1] dark:bg-black">
        <button
          type="button"
          onClick={() => setCollapsed(false)}
          aria-expanded={false}
          aria-label="Show uploads"
          title="Show uploads"
          className="rounded-full border border-black/[.1] p-2 text-xs font-medium text-zinc-600 transition-colors hover:bg-black/[.04] dark:border-white/[.15] dark:text-zinc-300 dark:hover:bg-white/[.06]"
        >
          ▸
        </button>
      </aside>
    );
  }

  return (
    <aside className="sticky top-0 flex h-screen w-72 shrink-0 flex-col overflow-y-auto border-r border-black/[.08] bg-white dark:border-white/[.1] dark:bg-black">
      <div className="flex items-center justify-between gap-2 border-b border-black/[.08] px-4 py-4 dark:border-white/[.1]">
        <span className="text-sm font-semibold text-black dark:text-zinc-50">
          Uploads
        </span>
        <button
          type="button"
          onClick={() => setCollapsed(true)}
          aria-expanded={true}
          aria-label="Hide uploads"
          title="Hide uploads"
          className="rounded-full border border-black/[.1] px-2 py-1 text-xs font-medium text-zinc-600 transition-colors hover:bg-black/[.04] dark:border-white/[.15] dark:text-zinc-300 dark:hover:bg-white/[.06]"
        >
          ◂
        </button>
      </div>

      <div className="flex flex-col gap-6 px-4 py-4">
        <section className="flex flex-col gap-3">
          <h2 className="text-xs font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400">
            Reports from system
          </h2>
          <div className="grid w-full grid-cols-1 gap-4">
            {PAYMENT_METHOD_ORDER.map((method) => (
              <FileDropzone
                key={method}
                label={`${PAYMENT_METHOD_LABELS[method]} report`}
                description="Filtered to this payment method"
                onFileParsed={onPaymentUpload(method)}
              />
            ))}
            <FileDropzone
              label="Expense report"
              description="From the system, all clinics (Account column determines payment method)"
              onFileParsed={onExpenseUpload}
            />
          </div>
        </section>

        <hr className="border-black/[.08] dark:border-white/[.1]" />

        <section className="flex flex-col gap-3">
          <h2 className="text-xs font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400">
            Manual staff reports
          </h2>
          <div className="grid w-full grid-cols-1 gap-4">
            {MANUAL_REPORT_CLINICS.map((clinic) => (
              <FileDropzone
                key={clinic.id}
                label={clinic.label}
                description="Manual report compiled by our staff"
                onFileParsed={onManualUpload(clinic.id)}
              />
            ))}
          </div>
        </section>
      </div>
    </aside>
  );
}
