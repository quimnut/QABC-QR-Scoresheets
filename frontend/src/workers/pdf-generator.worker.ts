/**
 * PDF Generator Web Worker
 *
 * Receives a 'generate' message, produces per-category PDFs using pdf-lib,
 * and streams progress back to the main thread.
 */

import { PDFDocument, PDFTextField, rgb } from "pdf-lib";
import type { GeneratorWorkerInput, GeneratorWorkerOutput, Entry } from "../lib/types";
import { getStyleSetById, getEntryType } from "../lib/style-sets";
import { makeQRBytes, recommendTextSize } from "../lib/qr";
import { sortCategoryEntries } from "../lib/pdf-generator";
import { padJudgingNumber } from "../lib/judging-number";

// ─── Constants ─────────────────────────────────────────────────────────────
const QR_RECT = { x: 205, y: 51.89, width: 70, height: 70 } as const;
const LOGO_RECT = { x: 529, y: 775.89, width: 60, height: 60 } as const;

// ─── Message handler ───────────────────────────────────────────────────────

self.onmessage = async (evt: MessageEvent<GeneratorWorkerInput>) => {
  const msg = evt.data;
  if (msg.type !== "generate") return;

  const { templateBytes, entries, logoBytes, copies, styleSetId, singlePdf } = msg;
  const styleSet = getStyleSetById(styleSetId);
  const logoBuf = logoBytes ? new Uint8Array(logoBytes) : null;

  try {
    // Group entries by category
    const byCategory = new Map<string, Entry[]>();
    for (const entry of entries) {
      const cat = String(entry.Category ?? "").trim();
      if (!cat || cat === "0") continue;
      if (!byCategory.has(cat)) byCategory.set(cat, []);
      byCategory.get(cat)!.push(entry);
    }

    const categories = [...byCategory.keys()].sort((a, b) => {
      const na = parseInt(a, 10);
      const nb = parseInt(b, 10);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a.localeCompare(b);
    });

    const totalEntries = entries.length;
    let completedEntries = 0;

    const pdfs: { name: string; bytes: ArrayBuffer }[] = [];

    // Single-PDF mode: one shared document for all categories
    const allDoc = singlePdf ? await PDFDocument.create() : null;

    for (const cat of categories) {
      const catEntries = sortCategoryEntries(byCategory.get(cat)!);
      const catNum = parseInt(cat, 10);
      const catLabel = isNaN(catNum) ? cat : String(catNum).padStart(2, "0");
      const outputDoc = allDoc ?? (await PDFDocument.create());

      for (const entry of catEntries) {
        const entryType = getEntryType(entry, styleSet);
        const templateBuf = templateBytes[entryType];

        const filledBytes = await fillEntry(templateBuf, entry, logoBuf);
        const filledDoc = await PDFDocument.load(filledBytes);

        for (let c = 0; c < copies; c++) {
          const [page] = await outputDoc.copyPages(filledDoc, [0]);
          outputDoc.addPage(page);
        }

        completedEntries++;
        const percent = Math.round((completedEntries / totalEntries) * 100);
        post({
          type: "progress",
          current: completedEntries,
          total: totalEntries,
          message: `Generating ${entryType} #${padJudgingNumber(entry["Judging Number"])} (cat ${catLabel})… ${percent}%`,
        });
      }

      if (!singlePdf && outputDoc.getPageCount() > 0) {
        const bytes = await outputDoc.save();
        pdfs.push({ name: `scoresheets_cat${catLabel}.pdf`, bytes: bytes.buffer as ArrayBuffer });
      }
    }

    if (singlePdf && allDoc && allDoc.getPageCount() > 0) {
      const bytes = await allDoc.save();
      pdfs.push({ name: "scoresheets_all.pdf", bytes: bytes.buffer as ArrayBuffer });
    }

    post({ type: "complete", pdfs });
  } catch (err) {
    post({ type: "error", message: String(err) });
  }
};

// ─── Fill one entry ────────────────────────────────────────────────────────

async function fillEntry(
  templateBuf: ArrayBuffer,
  entry: Entry,
  logoBytes: Uint8Array | null
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(templateBuf);
  const form = doc.getForm();
  const page = doc.getPages()[0];

  // QR code
  const qrBytes = await makeQRBytes(padJudgingNumber(entry["Judging Number"]));
  const qrImage = await doc.embedPng(qrBytes);
  page.drawImage(qrImage, QR_RECT);

  // Logo
  if (logoBytes) {
    page.drawRectangle({
      x: LOGO_RECT.x,
      y: LOGO_RECT.y,
      width: LOGO_RECT.width,
      height: LOGO_RECT.height,
      color: rgb(1, 1, 1),
    });
    const logoImage = await doc.embedPng(logoBytes);
    page.drawImage(logoImage, LOGO_RECT);
  }

  // Footer
  const filledEntry: Entry = {
    ...entry,
    FooterText: `Location : ${entry.Location} | Table : ${entry.Table} | Flight : ${entry.Flight}`,
  };

  // Fill AcroForm fields
  const specialFontSize = recommendTextSize(filledEntry.SpecialIngredients ?? "");
  for (const field of form.getFields()) {
    const name = field.getName();
    const value = filledEntry[name];
    if (value === undefined || value === null) continue;
    if (field instanceof PDFTextField) {
      const tf = form.getTextField(name);
      tf.setText(String(value));
      if (name === "SpecialIngredients") tf.setFontSize(specialFontSize);
    }
  }

  form.flatten();
  return doc.save();
}

// ─── Typed postMessage ─────────────────────────────────────────────────────

function post(msg: GeneratorWorkerOutput): void {
  self.postMessage(msg);
}
