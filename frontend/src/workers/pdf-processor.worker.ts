/**
 * PDF Processor Web Worker
 *
 * Receives a 'process' message with merged scanned PDF bytes, runs the
 * 5-strategy QR decode pipeline on each page, and returns sorted PDFs.
 */

import type { ProcessorWorkerInput, ProcessorWorkerOutput } from "../lib/types";
import { sortScannedScoresheets } from "../lib/pdf-processor";

// ─── Message handler ───────────────────────────────────────────────────────

self.onmessage = async (evt: MessageEvent<ProcessorWorkerInput>) => {
  const msg = evt.data;
  if (msg.type !== "process") return;

  try {
    const result = await sortScannedScoresheets(
      msg.pdfBytes,
      (done, total) => {
        post({
          type: "progress",
          current: done,
          total,
          message: `Scanning page ${done} of ${total}…`,
        });
      }
    );

    post({
      type: "complete",
      sorted: result.sorted.map((s) => ({
        name: s.name,
        bytes: s.bytes.buffer as ArrayBuffer,
      })),
      rejects: result.rejects ? (result.rejects.buffer as ArrayBuffer) : null,
      processed: result.processed,
      rejected: result.rejected,
    });
  } catch (err) {
    post({ type: "error", message: String(err) });
  }
};

// ─── Typed postMessage ─────────────────────────────────────────────────────

function post(msg: ProcessorWorkerOutput): void {
  self.postMessage(msg);
}
