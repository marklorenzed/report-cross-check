"use client";

import { useEffect, useState } from "react";
import {
  DaySummary,
  PAYMENT_METHOD_LABELS,
  PAYMENT_METHOD_ORDER,
  PaymentMethodTotals,
} from "../lib/collectionReport";
import { currency } from "../lib/currency";

type DayDetailDialogProps = {
  day: DaySummary;
  onClose: () => void;
};

type Tab = "clinic" | "dentist";

function BreakdownTable({
  labelHeader,
  rows,
  emptyMessage,
}: {
  labelHeader: string;
  rows: { label: string; totals: PaymentMethodTotals; total: number }[];
  emptyMessage: string;
}) {
  if (rows.length === 0) {
    return (
      <p className="text-xs text-zinc-500 dark:text-zinc-400">
        {emptyMessage}
      </p>
    );
  }

  return (
    <div className="overflow-x-auto">
      <table className="w-full min-w-[520px] text-xs">
        <thead>
          <tr className="text-zinc-500 dark:text-zinc-400">
            <th className="py-1.5 text-left font-medium">{labelHeader}</th>
            {PAYMENT_METHOD_ORDER.map((method) => (
              <th key={method} className="py-1.5 text-right font-medium">
                {PAYMENT_METHOD_LABELS[method]}
              </th>
            ))}
            <th className="py-1.5 text-right font-medium">Total</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr
              key={row.label}
              className="border-t border-black/[.06] transition-colors hover:bg-black/[.03] dark:border-white/[.08] dark:hover:bg-white/[.06]"
            >
              <td className="py-1.5 text-zinc-700 dark:text-zinc-300">
                {row.label}
              </td>
              {PAYMENT_METHOD_ORDER.map((method) => (
                <td key={method} className="py-1.5 text-right tabular-nums">
                  {row.totals[method] ? currency.format(row.totals[method]) : "—"}
                </td>
              ))}
              <td className="py-1.5 text-right font-semibold tabular-nums text-black dark:text-zinc-50">
                {currency.format(row.total)}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function DayDetailDialog({ day, onClose }: DayDetailDialogProps) {
  const [activeTab, setActiveTab] = useState<Tab>("clinic");

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onClose]);

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
      onClick={onClose}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-label={`Details for ${day.date}`}
        onClick={(event) => event.stopPropagation()}
        className="flex max-h-[80vh] w-full max-w-2xl flex-col gap-6 overflow-y-auto rounded-xl bg-white p-6 shadow-xl dark:bg-zinc-900"
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <h2 className="text-lg font-semibold text-black dark:text-zinc-50">
              {day.date}
            </h2>
            <p className="text-xs text-zinc-500 dark:text-zinc-400">
              Day total: {currency.format(day.total)}
            </p>
          </div>
          <button
            onClick={onClose}
            aria-label="Close"
            className="shrink-0 rounded-full p-1 text-zinc-500 transition-colors hover:bg-black/[.06] hover:text-black dark:text-zinc-400 dark:hover:bg-white/[.1] dark:hover:text-zinc-50"
          >
            ✕
          </button>
        </div>

        <div className="flex flex-col gap-3">
          <div
            role="tablist"
            className="flex w-fit gap-1 rounded-lg bg-black/[.04] p-1 dark:bg-white/[.06]"
          >
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === "clinic"}
              onClick={() => setActiveTab("clinic")}
              className={`rounded-md px-3 py-1.5 text-sm font-medium transition-colors ${
                activeTab === "clinic"
                  ? "bg-white text-black shadow-sm dark:bg-zinc-700 dark:text-zinc-50"
                  : "text-zinc-500 hover:text-black dark:text-zinc-400 dark:hover:text-zinc-50"
              }`}
            >
              By Clinic/Lab
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === "dentist"}
              onClick={() => setActiveTab("dentist")}
              className={`rounded-md px-3 py-1.5 text-sm font-medium transition-colors ${
                activeTab === "dentist"
                  ? "bg-white text-black shadow-sm dark:bg-zinc-700 dark:text-zinc-50"
                  : "text-zinc-500 hover:text-black dark:text-zinc-400 dark:hover:text-zinc-50"
              }`}
            >
              By Dentist
            </button>
          </div>

          {activeTab === "clinic" ? (
            <BreakdownTable
              labelHeader="Clinic/Lab"
              rows={day.clinics.map((clinic) => ({
                label: clinic.clinicLab,
                totals: clinic.totals,
                total: clinic.total,
              }))}
              emptyMessage="No clinic data for this day."
            />
          ) : (
            <BreakdownTable
              labelHeader="Dentist"
              rows={day.dentists.map((dentist) => ({
                label: dentist.name,
                totals: dentist.totals,
                total: dentist.total,
              }))}
              emptyMessage="No per-dentist sheets found in the uploaded reports."
            />
          )}
        </div>
      </div>
    </div>
  );
}
