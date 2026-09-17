"use client";

import { useCallback, useId, useRef, useState } from "react";
import * as XLSX from "xlsx";

export type ParsedWorkbook = {
  fileName: string;
  workbook: XLSX.WorkBook;
};

type FileDropzoneProps = {
  label: string;
  description?: string;
  onFileParsed: (data: ParsedWorkbook | null) => void;
};

const ACCEPTED_EXTENSIONS = [".xlsx", ".xls"];

function isExcelFile(file: File) {
  const name = file.name.toLowerCase();
  return ACCEPTED_EXTENSIONS.some((ext) => name.endsWith(ext));
}

export default function FileDropzone({
  label,
  description,
  onFileParsed,
}: FileDropzoneProps) {
  const inputId = useId();
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [rowCount, setRowCount] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);

  const parseFile = useCallback(
    async (file: File) => {
      setError(null);

      if (!isExcelFile(file)) {
        setError("Please select an .xlsx or .xls file.");
        setFileName(null);
        setRowCount(null);
        onFileParsed(null);
        return;
      }

      try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);

        setFileName(file.name);
        setRowCount(rows.length);
        onFileParsed({ fileName: file.name, workbook });
      } catch {
        setError("Could not read that file. Is it a valid Excel export?");
        setFileName(null);
        setRowCount(null);
        onFileParsed(null);
      }
    },
    [onFileParsed]
  );

  const handleDrop = useCallback(
    (event: React.DragEvent<HTMLLabelElement>) => {
      event.preventDefault();
      setIsDragActive(false);
      const file = event.dataTransfer.files?.[0];
      if (file) parseFile(file);
    },
    [parseFile]
  );

  const handleInputChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (file) parseFile(file);
    },
    [parseFile]
  );

  const handleClear = useCallback(
    (event: React.MouseEvent) => {
      event.preventDefault();
      event.stopPropagation();
      setFileName(null);
      setRowCount(null);
      setError(null);
      onFileParsed(null);
      if (inputRef.current) inputRef.current.value = "";
    },
    [onFileParsed]
  );

  return (
    <div className="flex w-full flex-col gap-2">
      <label
        htmlFor={inputId}
        onDragOver={(event) => {
          event.preventDefault();
          setIsDragActive(true);
        }}
        onDragLeave={() => setIsDragActive(false)}
        onDrop={handleDrop}
        className={`flex w-full cursor-pointer flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed px-6 py-10 text-center transition-colors ${
          isDragActive
            ? "border-foreground bg-black/[.04] dark:bg-white/[.08]"
            : "border-black/[.15] dark:border-white/[.2]"
        }`}
      >
        <span className="text-sm font-semibold text-black dark:text-zinc-50">
          {label}
        </span>
        {description && (
          <span className="text-xs text-zinc-500 dark:text-zinc-400">
            {description}
          </span>
        )}
        <span className="text-xs text-zinc-500 dark:text-zinc-400">
          Drag & drop an Excel file here, or click to browse
        </span>
        <input
          ref={inputRef}
          id={inputId}
          type="file"
          accept=".xlsx,.xls"
          className="sr-only"
          onChange={handleInputChange}
        />
      </label>

      {fileName && (
        <div className="flex items-center justify-between gap-2 rounded-lg bg-black/[.04] px-3 py-2 text-xs dark:bg-white/[.08]">
          <span className="truncate text-black dark:text-zinc-50">
            {fileName} · {rowCount} row{rowCount === 1 ? "" : "s"}
          </span>
          <button
            onClick={handleClear}
            className="shrink-0 font-medium text-zinc-500 hover:text-black dark:text-zinc-400 dark:hover:text-zinc-50"
          >
            Clear
          </button>
        </div>
      )}

      {error && (
        <span className="text-xs text-red-600 dark:text-red-400">
          {error}
        </span>
      )}
    </div>
  );
}
