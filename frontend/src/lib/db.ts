import { openDB, type DBSchema, type IDBPDatabase } from "idb";
import type { Entry, GeneratedPdf, ProcessedPdf, StyleSet, StyleSetConfidence, UploadedPdf } from "./types";

// ─── Schema ────────────────────────────────────────────────────────────────

interface QabcDB extends DBSchema {
  generate: {
    key: "current";
    value: GenerateStoreValue;
  };
  process: {
    key: "current";
    value: ProcessStoreValue;
  };
}

interface GenerateStoreValue {
  csvFilename: string | null;
  csvBytes: Uint8Array | null;
  entries: Entry[] | null;
  styleSet: StyleSet | null;
  styleSetId: number | null;
  styleSetConfidence: StyleSetConfidence | null;
  styleSetWarnings: string[];
  logoBytes: Uint8Array | null;
  generatedPdfs: { name: string; bytes: Uint8Array }[];
}

interface ProcessStoreValue {
  uploadedPdfs: {
    id: string;
    filename: string;
    bytes: Uint8Array;
    pages: number;
    processed: boolean;
  }[];
  processedPdfs: { name: string; bytes: Uint8Array }[];
  rejectsBytes: Uint8Array | null;
}

// ─── Default values ────────────────────────────────────────────────────────

const DEFAULT_GENERATE: GenerateStoreValue = {
  csvFilename: null,
  csvBytes: null,
  entries: null,
  styleSet: null,
  styleSetId: null,
  styleSetConfidence: null,
  styleSetWarnings: [],
  logoBytes: null,
  generatedPdfs: [],
};

const DEFAULT_PROCESS: ProcessStoreValue = {
  uploadedPdfs: [],
  processedPdfs: [],
  rejectsBytes: null,
};

// ─── DB singleton ──────────────────────────────────────────────────────────

let _db: IDBPDatabase<QabcDB> | null = null;

async function getDB(): Promise<IDBPDatabase<QabcDB>> {
  if (_db) return _db;
  _db = await openDB<QabcDB>("qabc-session", 1, {
    upgrade(db) {
      db.createObjectStore("generate");
      db.createObjectStore("process");
    },
  });
  return _db;
}

// ─── Generate store helpers ────────────────────────────────────────────────

export async function loadGenerateState(): Promise<GenerateStoreValue> {
  const db = await getDB();
  return (await db.get("generate", "current")) ?? DEFAULT_GENERATE;
}

export async function saveGenerateState(value: Partial<GenerateStoreValue>): Promise<void> {
  const db = await getDB();
  const current = (await db.get("generate", "current")) ?? DEFAULT_GENERATE;
  await db.put("generate", { ...current, ...value }, "current");
}

export async function clearGenerateState(): Promise<void> {
  const db = await getDB();
  await db.put("generate", DEFAULT_GENERATE, "current");
}

// ─── Process store helpers ─────────────────────────────────────────────────

export async function loadProcessState(): Promise<ProcessStoreValue> {
  const db = await getDB();
  return (await db.get("process", "current")) ?? DEFAULT_PROCESS;
}

export async function saveProcessState(value: Partial<ProcessStoreValue>): Promise<void> {
  const db = await getDB();
  const current = (await db.get("process", "current")) ?? DEFAULT_PROCESS;
  await db.put("process", { ...current, ...value }, "current");
}

export async function clearProcessState(): Promise<void> {
  const db = await getDB();
  await db.put("process", DEFAULT_PROCESS, "current");
}

// ─── Storage usage ─────────────────────────────────────────────────────────

/**
 * Returns the total bytes of all Uint8Array data stored in IndexedDB for this
 * session (uploaded PDFs, processed PDFs, rejects, generated PDFs, CSV, logo).
 * Works in both secure and non-secure contexts.
 */
export async function getStorageUsage(): Promise<number> {
  const db = await getDB();
  const gen = (await db.get("generate", "current")) ?? DEFAULT_GENERATE;
  const proc = (await db.get("process", "current")) ?? DEFAULT_PROCESS;

  let total = 0;
  total += gen.csvBytes?.length ?? 0;
  total += gen.logoBytes?.length ?? 0;
  for (const p of gen.generatedPdfs) total += p.bytes?.length ?? 0;
  for (const p of proc.uploadedPdfs) total += p.bytes?.length ?? 0;
  for (const p of proc.processedPdfs) total += p.bytes?.length ?? 0;
  total += proc.rejectsBytes?.length ?? 0;
  return total;
}

// ─── Re-export value types for consumers ──────────────────────────────────
export type { GenerateStoreValue, ProcessStoreValue };
