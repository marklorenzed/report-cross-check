"use client";

import { useMemo, useState } from "react";
import {
  ManualComparisonRow,
  PAYMENT_METHOD_LABELS,
  PAYMENT_METHOD_ORDER,
  PaymentMethodTotals,
} from "../lib/collectionReport";
import { currency } from "../lib/currency";

function totalOf(totals: PaymentMethodTotals): number {
  return PAYMENT_METHOD_ORDER.reduce((sum, method) => sum + totals[method], 0);
}

function isRowMatch(row: ManualComparisonRow): boolean {
  if (!row.reportTotals || !row.manualTotals) return false;
  return PAYMENT_METHOD_ORDER.every(
    (method) => row.reportTotals![method] === row.manualTotals![method]
  );
}

function ValueCell({
  manualValue,
  reportValue,
}: {
  manualValue: number | null;
  reportValue: number | null;
}) {
  if (manualValue !== null && reportValue !== null && manualValue !== reportValue) {
    return (
      <td className="px-3 py-1.5 text-right font-semibold tabular-nums text-red-600 dark:text-red-400">
        {currency.format(manualValue)} - {currency.format(reportValue)}
      </td>
    );
  }

  const value = manualValue ?? reportValue;
  return (
    <td className="px-3 py-1.5 text-right tabular-nums">
      {value !== null ? currency.format(value) : "—"}
    </td>
  );
}

export default function ManualComparisonTable({
  rows,
}: {
  rows: ManualComparisonRow[];
}) {
  const [clinicFilter, setClinicFilter] = useState("all");
  const [dateFilter, setDateFilter] = useState("all");
  const [mismatchesOnly, setMismatchesOnly] = useState(false);

  const nonZeroRows = useMemo(
    () =>
      rows.filter((row) => {
        const manualTotal = row.manualTotals ? totalOf(row.manualTotals) : 0;
        const reportTotal = row.reportTotals ? totalOf(row.reportTotals) : 0;
        return manualTotal !== 0 || reportTotal !== 0;
      }),
    [rows]
  );

  const clinicOptions = useMemo(
    () => [...new Set(nonZeroRows.map((row) => row.clinicLabel))],
    [nonZeroRows]
  );
  const dateOptions = useMemo(
    () => [...new Set(nonZeroRows.map((row) => row.date))],
    [nonZeroRows]
  );

  const visibleRows = nonZeroRows.filter(
    (row) =>
      (clinicFilter === "all" || row.clinicLabel === clinicFilter) &&
      (dateFilter === "all" || row.date === dateFilter) &&
      (!mismatchesOnly || !isRowMatch(row))
  );

  if (nonZeroRows.length === 0) {
    return (
      <p className="text-sm text-zinc-500 dark:text-zinc-400">
        Upload the payment method reports and the manual staff reports to
        compare them.
      </p>
    );
  }

  return (
    <div className="flex w-full flex-col gap-3">
      <div className="flex flex-wrap items-center gap-3 text-xs">
        <label className="flex items-center gap-2">
          <span className="text-zinc-500 dark:text-zinc-400">Date</span>
          <select
            value={dateFilter}
            onChange={(event) => setDateFilter(event.target.value)}
            className="rounded-md border border-black/[.1] bg-white px-2 py-1 text-black dark:border-white/[.15] dark:bg-zinc-900 dark:text-zinc-50"
          >
            <option value="all">All dates</option>
            {dateOptions.map((date) => (
              <option key={date} value={date}>
                {date}
              </option>
            ))}
          </select>
        </label>
        <label className="flex items-center gap-2">
          <span className="text-zinc-500 dark:text-zinc-400">Clinic</span>
          <select
            value={clinicFilter}
            onChange={(event) => setClinicFilter(event.target.value)}
            className="rounded-md border border-black/[.1] bg-white px-2 py-1 text-black dark:border-white/[.15] dark:bg-zinc-900 dark:text-zinc-50"
          >
            <option value="all">All clinics</option>
            {clinicOptions.map((clinic) => (
              <option key={clinic} value={clinic}>
                {clinic}
              </option>
            ))}
          </select>
        </label>
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={mismatchesOnly}
            onChange={(event) => setMismatchesOnly(event.target.checked)}
            className="h-3.5 w-3.5"
          />
          <span className="text-zinc-500 dark:text-zinc-400">
            Show only mismatches
          </span>
        </label>
      </div>

      {visibleRows.length === 0 ? (
        <p className="text-sm text-zinc-500 dark:text-zinc-400">
          No rows match the selected filters.
        </p>
      ) : (
        <div className="w-full overflow-x-auto rounded-lg border border-black/[.08] dark:border-white/[.1]">
          <table className="w-full min-w-[880px] text-xs">
            <thead>
              <tr className="bg-black/[.03] text-zinc-500 dark:bg-white/[.06] dark:text-zinc-400">
                <th className="px-3 py-2 text-left font-medium">Date</th>
                <th className="px-3 py-2 text-left font-medium">Clinic</th>
                {PAYMENT_METHOD_ORDER.map((method) => (
                  <th key={method} className="px-3 py-2 text-right font-medium">
                    {PAYMENT_METHOD_LABELS[method]}
                  </th>
                ))}
                <th className="px-3 py-2 text-right font-medium">Total</th>
                <th className="px-3 py-2 text-center font-medium">Match</th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((row) => {
                const match = isRowMatch(row);
                const manualTotal = row.manualTotals ? totalOf(row.manualTotals) : null;
                const reportTotal = row.reportTotals ? totalOf(row.reportTotals) : null;

                return (
                  <tr
                    key={`${row.date}__${row.clinicLabel}`}
                    className="border-t border-black/[.06] dark:border-white/[.08]"
                  >
                    <td className="px-3 py-1.5 font-medium text-black dark:text-zinc-50">
                      {row.date}
                    </td>
                    <td className="px-3 py-1.5 text-zinc-700 dark:text-zinc-300">
                      {row.clinicLabel}
                    </td>
                    {PAYMENT_METHOD_ORDER.map((method) => (
                      <ValueCell
                        key={method}
                        manualValue={row.manualTotals?.[method] ?? null}
                        reportValue={row.reportTotals?.[method] ?? null}
                      />
                    ))}
                    <ValueCell manualValue={manualTotal} reportValue={reportTotal} />
                    <td className="px-3 py-1.5 text-center">
                      {row.reportTotals && row.manualTotals ? (match ? "✅" : "❌") : "—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
