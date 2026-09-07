import { PDFDocument } from "pdf-lib";
import jsQR from "jsqr";
import { initPdfWorker } from "./pdfjs-init";
import { padJudgingNumber } from "./judging-number";

// ─── Types ────────────────────────────────────────────────────────────────

export interface SortResult {
  sorted: { name: string; bytes: Uint8Array }[];
  rejects: Uint8Array | null;
  processed: number;
  rejected: number;
}

// ─── Main processing entry point ──────────────────────────────────────────

/**
 * Sort a scanned PDF by QR code content.
 *
 * For each page:
 *  - Render to canvas (PDF.js)
 *  - Try 5 decoding strategies
 *  - Route to per-judging-number PDF or rejects
 *
 * @param pdfBytes  Merged scanned PDF as ArrayBuffer
 * @param onProgress Called after each page: (pagesDone, totalPages)
 * @param cancelCheck Called between pages; return true to abort
 */
export async function sortScannedScoresheets(
  pdfBytes: ArrayBuffer,
  onProgress?: (done: number, total: number) => void | Promise<void>,
  cancelCheck?: () => boolean
): Promise<SortResult> {
  const { getDocument, GlobalWorkerOptions } = await import("pdfjs-dist");
  await initPdfWorker(GlobalWorkerOptions);

  // pdfjs transfers the ArrayBuffer to its worker (detaching the original).
  // Save a copy for pdf-lib operations that happen after rendering.
  const pdfBytesForLib = pdfBytes.slice(0);

  const loadingTask = getDocument({ data: new Uint8Array(pdfBytes) });
  const pdfDoc = await loadingTask.promise;
  const numPages = pdfDoc.numPages;

  // Accumulate pages per judging number
  const pageMap = new Map<string, number[]>(); // judging number → page indices
  const rejectPages: number[] = [];

  for (let i = 0; i < numPages; i++) {
    if (cancelCheck?.()) break;

    const page = await pdfDoc.getPage(i + 1); // pdf.js is 1-indexed
    const imageData = await renderPageToImageData(page, 1.0);

    const decoded = decodeQR(imageData);
    if (decoded) {
      const padded = padJudgingNumber(decoded);
      if (!pageMap.has(padded)) pageMap.set(padded, []);
      pageMap.get(padded)!.push(i);
    } else {
      rejectPages.push(i);
    }

    onProgress?.(i + 1, numPages);
    // Yield to the event loop so the caller can update UI between pages
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  }

  loadingTask.destroy();

  // Build sorted output PDFs
  const sorted: { name: string; bytes: Uint8Array }[] = [];
  for (const [judgingNumber, pages] of pageMap) {
    const outDoc = await PDFDocument.create();
    const srcDoc = await PDFDocument.load(pdfBytesForLib);
    for (const pageIdx of pages) {
      const [copied] = await outDoc.copyPages(srcDoc, [pageIdx]);
      outDoc.addPage(copied);
    }
    sorted.push({ name: `${padJudgingNumber(judgingNumber)}.pdf`, bytes: await outDoc.save() });
  }

  // Build rejects PDF
  let rejectsBytes: Uint8Array | null = null;
  if (rejectPages.length > 0) {
    const rejectDoc = await PDFDocument.create();
    const srcDoc = await PDFDocument.load(pdfBytesForLib);
    for (const pageIdx of rejectPages) {
      const [copied] = await rejectDoc.copyPages(srcDoc, [pageIdx]);
      rejectDoc.addPage(copied);
    }
    rejectsBytes = await rejectDoc.save();
  }

  return {
    sorted,
    rejects: rejectsBytes,
    processed: numPages - rejectPages.length,
    rejected: rejectPages.length,
  };
}

// ─── PDF page operations ──────────────────────────────────────────────────

/** Get the number of pages in a PDF Uint8Array. */
export async function getPdfPageCount(bytes: Uint8Array): Promise<number> {
  const doc = await PDFDocument.load(bytes);
  return doc.getPageCount();
}

/** Delete a page (0-indexed) from a PDF. Returns null if no pages remain. */
export async function deletePage(bytes: Uint8Array, pageIndex: number): Promise<Uint8Array | null> {
  const doc = await PDFDocument.load(bytes);
  if (doc.getPageCount() <= 1) return null;
  doc.removePage(pageIndex);
  return doc.save();
}

/**
 * Move a page from `rejectsBytes` (0-indexed) into the correct sorted PDF.
 * Returns { updatedRejects, updatedTarget }.
 */
export async function assignRejectPage(
  rejectsBytes: Uint8Array,
  pageIndex: number,
  judgingNumber: string,
  existingTargetBytes: Uint8Array | null
): Promise<{ updatedRejects: Uint8Array | null; updatedTarget: Uint8Array }> {
  // Pad judging number to 6 digits
  const paddedName = padJudgingNumber(judgingNumber);

  // Copy the reject page into the target PDF
  const rejectsDoc = await PDFDocument.load(rejectsBytes);
  const targetDoc = existingTargetBytes
    ? await PDFDocument.load(existingTargetBytes)
    : await PDFDocument.create();

  const [copiedPage] = await targetDoc.copyPages(rejectsDoc, [pageIndex]);
  targetDoc.addPage(copiedPage);
  const updatedTarget = await targetDoc.save();

  // Remove the page from rejects
  const updatedRejects = await deletePage(rejectsBytes, pageIndex);

  return { updatedRejects, updatedTarget };
}

/** Merge multiple PDFs into one. */
export async function mergePdfs(pdfList: Uint8Array[]): Promise<Uint8Array> {
  const mergedDoc = await PDFDocument.create();
  for (const bytes of pdfList) {
    const srcDoc = await PDFDocument.load(bytes);
    const pages = await mergedDoc.copyPages(srcDoc, srcDoc.getPageIndices());
    pages.forEach((p) => mergedDoc.addPage(p));
  }
  return mergedDoc.save();
}

// ─── Page rendering (for UI preview) ─────────────────────────────────────

/**
 * Render a single page of a PDF Uint8Array to a PNG data URL.
 * Uses PDF.js. Main-thread only (requires Canvas).
 */
export async function renderPdfPageToDataUrl(
  bytes: Uint8Array,
  pageIndex = 0,
  scale = 1.5
): Promise<string> {
  const { getDocument, GlobalWorkerOptions } = await import("pdfjs-dist");
  await initPdfWorker(GlobalWorkerOptions);

  // pdfjs transfers the ArrayBuffer to its worker — pass a copy so the
  // caller's Uint8Array is not detached (neutered) after this call.
  const loadingTask = getDocument({ data: new Uint8Array(bytes) });
  const pdfDoc = await loadingTask.promise;
  const page = await pdfDoc.getPage(pageIndex + 1);
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement("canvas");
  canvas.width = Math.round(viewport.width);
  canvas.height = Math.round(viewport.height);
  const ctx = canvas.getContext("2d")!;
  await page.render({ canvas, canvasContext: ctx, viewport }).promise;
  loadingTask.destroy();

  return canvas.toDataURL("image/png");
}

// ─── Internal: render page to ImageData ──────────────────────────────────

async function renderPageToImageData(
  page: Awaited<ReturnType<Awaited<ReturnType<typeof import("pdfjs-dist")["getDocument"]>["promise"]>["getPage"]>>,
  scale: number
): Promise<ImageData> {
  const viewport = page.getViewport({ scale });

  // OffscreenCanvas is available in Web Workers and modern browsers
  const canvas = new OffscreenCanvas(Math.round(viewport.width), Math.round(viewport.height));
  const ctx = canvas.getContext("2d") as OffscreenCanvasRenderingContext2D;

  await page.render({
    canvas: canvas as unknown as HTMLCanvasElement,
    canvasContext: ctx as unknown as CanvasRenderingContext2D,
    viewport,
  }).promise;

  return ctx.getImageData(0, 0, canvas.width, canvas.height);
}

// ─── 5-strategy QR decoding pipeline ─────────────────────────────────────

function decodeQR(imageData: ImageData): string | null {
  const { data, width, height } = imageData;

  // Strategy 1 — full RGBA
  const r1 = jsQR(data, width, height, { inversionAttempts: "dontInvert" });
  if (r1) return r1.data;

  // Strategy 2 — bottom-centre crop (x: 20%–60%, y: 80%–100%)
  const crop = cropImageData(imageData, 0.2, 0.8, 0.8, 1.0);
  const r2 = jsQR(crop.data, crop.width, crop.height, { inversionAttempts: "dontInvert" });
  if (r2) return r2.data;

  // Strategy 3 — adaptive threshold on full image, then jsQR
  const thresh = adaptiveThreshold(imageData);
  const r3 = jsQR(thresh.data, thresh.width, thresh.height, { inversionAttempts: "dontInvert" });
  if (r3) return r3.data;

  // Strategy 4 — try inverted colours
  const r4 = jsQR(data, width, height, { inversionAttempts: "invertFirst" });
  if (r4) return r4.data;

  // Strategy 5 — rotation sweep on crop (90°, 180°, 270° first, then multiples of 45°)
  const angles = [90, 180, 270, 45, 135, 225, 315, 30, 60, 120, 150, 210, 240, 300, 330];
  for (const angle of angles) {
    const rotated = rotateImageData(crop, angle);
    const rThresh = adaptiveThreshold(rotated);
    const r5 = jsQR(rThresh.data, rThresh.width, rThresh.height, { inversionAttempts: "dontInvert" });
    if (r5) return r5.data;
  }

  return null;
}

// ─── Image processing utilities ───────────────────────────────────────────

/** Crop a region from ImageData as fractions of width/height. */
function cropImageData(
  img: ImageData,
  x0: number,
  x1: number,
  y0: number,
  y1: number
): ImageData {
  const sx = Math.floor(img.width * x0);
  const ex = Math.floor(img.width * x1);
  const sy = Math.floor(img.height * y0);
  const ey = Math.floor(img.height * y1);
  const cw = ex - sx;
  const ch = ey - sy;
  const out = new Uint8ClampedArray(cw * ch * 4);
  for (let y = 0; y < ch; y++) {
    for (let x = 0; x < cw; x++) {
      const si = ((sy + y) * img.width + (sx + x)) * 4;
      const di = (y * cw + x) * 4;
      out[di] = img.data[si];
      out[di + 1] = img.data[si + 1];
      out[di + 2] = img.data[si + 2];
      out[di + 3] = img.data[si + 3];
    }
  }
  return new ImageData(out, cw, ch);
}

/**
 * Adaptive threshold (local mean, block 11×11, constant 2).
 * Uses integral image for O(1) per-pixel lookup.
 * Input: RGBA ImageData. Output: RGBA binary ImageData (0 or 255).
 */
function adaptiveThreshold(img: ImageData, blockSize = 11, c = 2): ImageData {
  const { width, height, data } = img;
  const gray = new Float32Array(width * height);
  for (let i = 0; i < gray.length; i++) {
    gray[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];
  }

  // Build integral image (padded by 1 on top/left)
  const iw = width + 1;
  const ih = height + 1;
  const integral = new Float64Array(iw * ih);
  for (let y = 1; y < ih; y++) {
    for (let x = 1; x < iw; x++) {
      integral[y * iw + x] =
        gray[(y - 1) * width + (x - 1)] +
        integral[(y - 1) * iw + x] +
        integral[y * iw + (x - 1)] -
        integral[(y - 1) * iw + (x - 1)];
    }
  }

  const half = Math.floor(blockSize / 2);
  const out = new Uint8ClampedArray(width * height * 4);

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const x1 = Math.max(0, x - half);
      const x2 = Math.min(width - 1, x + half);
      const y1 = Math.max(0, y - half);
      const y2 = Math.min(height - 1, y + half);
      const count = (x2 - x1 + 1) * (y2 - y1 + 1);
      const sum =
        integral[(y2 + 1) * iw + (x2 + 1)] -
        integral[y1 * iw + (x2 + 1)] -
        integral[(y2 + 1) * iw + x1] +
        integral[y1 * iw + x1];
      const mean = sum / count;
      const val = gray[y * width + x] < mean - c ? 0 : 255;
      const idx = (y * width + x) * 4;
      out[idx] = out[idx + 1] = out[idx + 2] = val;
      out[idx + 3] = 255;
    }
  }

  return new ImageData(out, width, height);
}

/** Rotate ImageData by angle degrees using OffscreenCanvas. */
function rotateImageData(img: ImageData, angle: number): ImageData {
  const radians = (angle * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const newW = Math.round(img.width * cos + img.height * sin);
  const newH = Math.round(img.width * sin + img.height * cos);

  const src = new OffscreenCanvas(img.width, img.height);
  const srcCtx = src.getContext("2d") as OffscreenCanvasRenderingContext2D;
  srcCtx.putImageData(img, 0, 0);

  const dst = new OffscreenCanvas(newW, newH);
  const dstCtx = dst.getContext("2d") as OffscreenCanvasRenderingContext2D;
  dstCtx.translate(newW / 2, newH / 2);
  dstCtx.rotate(radians);
  dstCtx.drawImage(src, -img.width / 2, -img.height / 2);

  return dstCtx.getImageData(0, 0, newW, newH);
}
