"use client";

import { useMemo, useState } from "react";
import type * as XLSX from "xlsx";
import FileDropzone, { ParsedWorkbook } from "./components/FileDropzone";
import CollectionSummary from "./components/CollectionSummary";
import {
  PAYMENT_METHOD_LABELS,
  PAYMENT_METHOD_ORDER,
  PaymentMethod,
  summarizeCollectionUploads,
} from "./lib/collectionReport";

export default function Home() {
  const [paymentUploads, setPaymentUploads] = useState<
    Partial<Record<PaymentMethod, ParsedWorkbook>>
  >({});
  const [manualReport, setManualReport] = useState<ParsedWorkbook | null>(
    null
  );

  const collectionDays = useMemo(() => {
    const workbooksByMethod: Partial<Record<PaymentMethod, XLSX.WorkBook>> = {};
    for (const method of PAYMENT_METHOD_ORDER) {
      const upload = paymentUploads[method];
      if (upload) workbooksByMethod[method] = upload.workbook;
    }
    if (Object.keys(workbooksByMethod).length === 0) return null;
    return summarizeCollectionUploads(workbooksByMethod);
  }, [paymentUploads]);

  const canCompare = collectionDays !== null && manualReport !== null;

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
            label="Manual staff report"
            description="Compiled by our staff"
            onFileParsed={setManualReport}
          />
        </div>

        <button
          disabled={!canCompare}
          className="flex h-12 w-full max-w-xs items-center justify-center rounded-full bg-foreground px-5 text-sm font-medium text-background transition-colors enabled:hover:bg-[#383838] disabled:cursor-not-allowed disabled:opacity-40 dark:enabled:hover:bg-[#ccc]"
        >
          Cross-check reports
        </button>

        {collectionDays && (
          <div className="flex w-full flex-col gap-4">
            <h2 className="text-lg font-semibold text-black dark:text-zinc-50">
              Collection Report summary
            </h2>
            <CollectionSummary days={collectionDays} />
          </div>
        )}
      </main>
    </div>
  );
}

