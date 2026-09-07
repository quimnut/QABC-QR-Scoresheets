import Papa from "papaparse";
import type { Entry, StyleSetDetectionResult } from "./types";
import { detectStyleSet } from "./style-sets";

// ─── CSV parsing ───────────────────────────────────────────────────────────

/**
 * Parse a CSV file (BCOEM format, latin-1 encoded) and return normalised entries.
 * Handles encoding in the browser by decoding via TextDecoder with 'latin1'.
 */
export function parseCSV(buffer: ArrayBuffer): Entry[] {
  // BCOEM exports are latin-1 encoded
  const text = new TextDecoder("latin1").decode(buffer);
  const result = Papa.parse<Record<string, string>>(text, {
    header: true,
    skipEmptyLines: true,
    transformHeader: (h: string) => h.trim(),
  });

  return (result.data as Record<string, string>[]).map((row) => {
    const specialIngredients = (row["Required Info"] ?? "")
      .replace(/\n/g, "; ")
      .trim();

    return {
      ...row,
      SubCategory: row["Subcategory"] ?? "",
      EntryNumber: row["Judging Number"] ?? "",
      SubCategoryName: row["Style"] ?? "",
      SpecialIngredients: specialIngredients,
      // ensure required keys exist with fallbacks
      "Judging Number": row["Judging Number"] ?? "",
      Location: row["Location"] ?? "",
      Table: row["Table"] ?? "",
      Flight: row["Flight"] ?? "",
      Round: row["Round"] ?? "",
      Style: row["Style"] ?? "",
      Category: row["Category"] ?? "",
      Subcategory: row["Subcategory"] ?? "",
      "Required Info": row["Required Info"] ?? "",
    } as Entry;
  });
}

// ─── Preview helpers ───────────────────────────────────────────────────────

export const PREVIEW_HEADERS = [
  "Judging Number",
  "Style",
  "Required Info",
  "Location",
  "Table",
  "Flight",
  "Round",
] as const;

export function getEntriesPreview(entries: Entry[]): { headers: string[]; rows: string[][] } {
  const headers = [...PREVIEW_HEADERS];
  const rows = entries.map((e) => headers.map((h) => String(e[h] ?? "")));
  return { headers, rows };
}

export function getUniqueCategories(entries: Entry[]): number[] {
  const cats = new Set<number>();
  for (const e of entries) {
    const n = parseInt(String(e.Category ?? ""), 10);
    if (!isNaN(n) && n > 0) cats.add(n);
  }
  return [...cats].sort((a, b) => a - b);
}

export function filterEntriesByCategory(entries: Entry[], category: number): Entry[] {
  return entries.filter((e) => parseInt(String(e.Category ?? ""), 10) === category);
}

// ─── Style set detection wrapper ──────────────────────────────────────────

export function detectAndStoreStyleSet(entries: Entry[]): StyleSetDetectionResult {
  return detectStyleSet(entries);
}

// ─── Dummy entry for template preview ─────────────────────────────────────

export function createDummyEntry(): Entry {
  return {
    "Judging Number": "SAMPLE",
    Location: "Sample Day 1",
    Table: "09: Table 9: Stouts",
    Flight: "1",
    Round: "1",
    Style: "Imperial Stout [BJCP 20C]",
    SpecialIngredients: "Sample Special Ingredients",
    SubCategory: "04",
    Category: "10",
    EntryNumber: "SAMPLE",
    SubCategoryName: "Imperial Stout [BJCP 20C]",
    Subcategory: "04",
    "Required Info": "Sample Special Ingredients",
    "Style Type": "beer",
  };
}

// ─── Format file size ──────────────────────────────────────────────────────

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
