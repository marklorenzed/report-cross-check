"use client";

import { useState } from "react";
import {
  DaySummary,
  PAYMENT_METHOD_LABELS,
  PAYMENT_METHOD_ORDER,
  PaymentMethodTotals,
} from "../lib/collectionReport";
import { currency } from "../lib/currency";
import DayDetailDialog from "./DayDetailDialog";

function TotalsCells({
  totals,
  total,
}: {
  totals: PaymentMethodTotals;
  total: number;
}) {
  return (
    <>
      {PAYMENT_METHOD_ORDER.map((method) => (
        <td key={method} className="px-3 py-2 text-right tabular-nums">
          {totals[method] ? currency.format(totals[method]) : "—"}
        </td>
      ))}
      <td className="px-3 py-2 text-right font-semibold tabular-nums">
        {currency.format(total)}
      </td>
    </>
  );
}

function DayRows({
  day,
  onSelect,
}: {
  day: DaySummary;
  onSelect: (day: DaySummary) => void;
}) {
  return (
    <>
      {day.clinics.map((clinic, index) => (
        <tr
          key={clinic.clinicLab}
          className="border-t border-black/[.06] transition-colors hover:bg-black/[.02] dark:border-white/[.08] dark:hover:bg-white/[.04]"
        >
          {index === 0 && (
            <td
              rowSpan={day.clinics.length + 1}
              role="button"
              tabIndex={0}
              onClick={() => onSelect(day)}
              onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === " ") {
                  event.preventDefault();
                  onSelect(day);
                }
              }}
              className="cursor-pointer px-3 py-2 align-top font-medium text-black transition-colors hover:bg-black/[.06] hover:underline focus:outline-none focus-visible:underline dark:text-zinc-50 dark:hover:bg-white/[.1]"
            >
              {day.date}
            </td>
          )}
          <td className="px-3 py-2 text-zinc-700 dark:text-zinc-300">
            {clinic.clinicLab}
          </td>
          <TotalsCells totals={clinic.totals} total={clinic.total} />
        </tr>
      ))}
      <tr className="border-t border-black/[.08] bg-black/[.03] font-semibold dark:border-white/[.1] dark:bg-white/[.06]">
        <td className="px-3 py-2 text-black dark:text-zinc-50">Day total</td>
        <TotalsCells totals={day.totals} total={day.total} />
      </tr>
    </>
  );
}

export default function CollectionSummary({ days }: { days: DaySummary[] }) {
  const [selectedDay, setSelectedDay] = useState<DaySummary | null>(null);

  if (days.length === 0) {
    return (
      <p className="text-sm text-zinc-500 dark:text-zinc-400">
        No rows found yet. Upload at least one payment method report.
      </p>
    );
  }

  return (
    <div className="w-full overflow-x-auto rounded-lg border border-black/[.08] dark:border-white/[.1]">
      <table className="w-full min-w-[720px] text-xs">
        <thead>
          <tr className="bg-black/[.03] text-zinc-500 dark:bg-white/[.06] dark:text-zinc-400">
            <th className="px-3 py-2 text-left font-medium">Date</th>
            <th className="px-3 py-2 text-left font-medium">Clinic/Lab</th>
            {PAYMENT_METHOD_ORDER.map((method) => (
              <th key={method} className="px-3 py-2 text-right font-medium">
                {PAYMENT_METHOD_LABELS[method]}
              </th>
            ))}
            <th className="px-3 py-2 text-right font-medium">Total</th>
          </tr>
        </thead>
        <tbody>
          {days.map((day) => (
            <DayRows key={day.date} day={day} onSelect={setSelectedDay} />
          ))}
        </tbody>
      </table>

      {selectedDay && (
        <DayDetailDialog day={selectedDay} onClose={() => setSelectedDay(null)} />
      )}
    </div>
  );
}
