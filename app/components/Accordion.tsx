"use client";

import { useState, type ReactNode } from "react";

export default function Accordion({
  title,
  defaultOpen = true,
  children,
}: {
  title: string;
  defaultOpen?: boolean;
  children: ReactNode;
}) {
  const [open, setOpen] = useState(defaultOpen);

  return (
    <div className="flex w-full flex-col gap-4">
      <button
        type="button"
        onClick={() => setOpen((prev) => !prev)}
        aria-expanded={open}
        className="flex w-full items-center justify-between gap-2 text-left"
      >
        <h2 className="text-lg font-semibold text-black dark:text-zinc-50">
          {title}
        </h2>
        <span
          className={`text-zinc-500 transition-transform dark:text-zinc-400 ${
            open ? "rotate-180" : ""
          }`}
        >
          ▾
        </span>
      </button>
      {open && children}
    </div>
  );
}
