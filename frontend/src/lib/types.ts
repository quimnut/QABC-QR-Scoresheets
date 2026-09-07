// ─── Core data types ───────────────────────────────────────────────────────

export interface Entry {
  "Judging Number": string;
  Location: string;
  Table: string;
  Flight: string;
  Round: string;
  Style: string;
  Category: string;
  Subcategory: string;
  "Required Info": string;
  "Style Type"?: string;
  // Synthesised fields added by csv-parser
  SubCategory: string;
  EntryNumber: string;
  SubCategoryName: string;
  SpecialIngredients: string;
  // Synthesised during generation
  FooterText?: string;
  [key: string]: string | undefined;
}

export interface StyleSet {
  id: number;
  name: string;
  long_name: string;
  short_name: string;
  beer_end: number;
  mead: string[];
  cider: string[];
  categories: Record<string, string>;
}

export type EntryType = "beer" | "mead" | "cider";
export type StyleSetConfidence = "high" | "medium" | "low";

export interface StyleSetDetectionResult {
  styleSet: StyleSet;
  confidence: StyleSetConfidence;
  warnings: string[];
  name: string;
}

// ─── Session / storage types ───────────────────────────────────────────────

export interface GeneratedPdf {
  name: string;
  bytes: Uint8Array;
}

export interface UploadedPdf {
  id: string;
  filename: string;
  bytes: Uint8Array;
  pages: number;
  processed: boolean;
}

export interface ProcessedPdf {
  name: string;
  bytes: Uint8Array;
}

// ─── Progress / worker message types ──────────────────────────────────────

export interface ProgressState {
  current: number;
  total: number;
  percent: number;
  message: string;
  complete: boolean;
}

// Messages sent TO the generator worker
export type GeneratorWorkerInput =
  | {
      type: "generate";
      templateBytes: { beer: ArrayBuffer; cider: ArrayBuffer; mead: ArrayBuffer };
      entries: Entry[];
      logoBytes: ArrayBuffer | null;
      copies: number;
      styleSetId: number;
      singlePdf: boolean;
    }
  | { type: "cancel" };

// Messages sent FROM the generator worker
export type GeneratorWorkerOutput =
  | { type: "progress"; current: number; total: number; message: string }
  | { type: "complete"; pdfs: { name: string; bytes: ArrayBuffer }[] }
  | { type: "error"; message: string };

// Messages sent TO the processor worker
export type ProcessorWorkerInput =
  | { type: "process"; pdfBytes: ArrayBuffer }
  | { type: "cancel" };

// Messages sent FROM the processor worker
export type ProcessorWorkerOutput =
  | { type: "progress"; current: number; total: number; message: string }
  | {
      type: "complete";
      sorted: { name: string; bytes: ArrayBuffer }[];
      rejects: ArrayBuffer | null;
      processed: number;
      rejected: number;
    }
  | { type: "error"; message: string };

// ─── UI types ─────────────────────────────────────────────────────────────

export interface Alert {
  id: string;
  kind: "success" | "error" | "warning" | "info";
  message: string;
}

export type ModalContent =
  | { kind: "image"; src: string; title: string }
  | { kind: "pdf"; bytes: Uint8Array; title: string }
  | { kind: "loading"; title: string };
