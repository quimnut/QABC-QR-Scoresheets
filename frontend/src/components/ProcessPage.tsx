import { useState, useEffect, useRef, useCallback } from "react";
import type { Alert, ModalContent } from "../lib/types";
import {
  loadProcessState,
  saveProcessState,
  clearProcessState,
  getStorageUsage,
} from "../lib/db";
import {
  mergePdfs,
  getPdfPageCount,
  deletePage,
  assignRejectPage,
  renderPdfPageToDataUrl,
  sortScannedScoresheets,
} from "../lib/pdf-processor";
import { formatFileSize } from "../lib/csv-parser";
import { padJudgingNumber } from "../lib/judging-number";
import AlertMessage from "./AlertMessage";
import ProgressBar from "./ProgressBar";
import Modal from "./Modal";
import JSZip from "jszip";

// ─── Local types ──────────────────────────────────────────────────────────

interface UploadedPdf {
  id: string;
  filename: string;
  bytes: Uint8Array;
  pages: number;
  processed: boolean;
}

interface ProcessedPdf {
  name: string;
  bytes: Uint8Array;
}

// ─── Component ─────────────────────────────────────────────────────────────

export default function ProcessPage() {
  // ── State ────────────────────────────────────────────────────────────────
  const [uploadedPdfs, setUploadedPdfs] = useState<UploadedPdf[]>([]);
  const [processedPdfs, setProcessedPdfs] = useState<ProcessedPdf[]>([]);
  const [rejectsBytes, setRejectsBytes] = useState<Uint8Array | null>(null);
  const [rejectPage, setRejectPage] = useState(0); // 0-indexed current page
  const [rejectPageCount, setRejectPageCount] = useState(0);
  const [rejectDataUrl, setRejectDataUrl] = useState<string | null>(null);
  const [rejectLoadingPage, setRejectLoadingPage] = useState(false);
  const [assignInput, setAssignInput] = useState("");
  const [progress, setProgress] = useState<{ current: number; total: number; message: string } | null>(null);
  const [processing, setProcessing] = useState(false);
  const [alerts, setAlerts] = useState<Alert[]>([]);
  const [modal, setModal] = useState<ModalContent | null>(null);
  const [storageUsage, setStorageUsage] = useState<number>(0);

  const cancelRef = useRef(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Alerts ───────────────────────────────────────────────────────────────
  const addAlert = useCallback((kind: Alert["kind"], message: string) => {
    const id = `${Date.now()}-${Math.random()}`;
    setAlerts((prev) => [...prev, { id, kind, message }]);
  }, []);

  const dismissAlert = useCallback((id: string) => {
    setAlerts((prev) => prev.filter((a) => a.id !== id));
  }, []);

  const refreshStorage = useCallback(async () => {
    setStorageUsage(await getStorageUsage());
  }, []);

  // ── Load persisted state on mount ────────────────────────────────────────
  useEffect(() => {
    loadProcessState().then((state) => {
      if (state.uploadedPdfs?.length) setUploadedPdfs(state.uploadedPdfs);
      if (state.processedPdfs?.length) setProcessedPdfs(state.processedPdfs);
      if (state.rejectsBytes) {
        setRejectsBytes(state.rejectsBytes);
        getPdfPageCount(state.rejectsBytes).then(setRejectPageCount);
      }
      refreshStorage();
    });
  }, [refreshStorage]);

  // ── Render current reject page whenever it or the bytes changes ───────────
  useEffect(() => {
    if (!rejectsBytes || rejectPageCount === 0) {
      setRejectDataUrl(null);
      return;
    }
    setRejectLoadingPage(true);
    renderPdfPageToDataUrl(rejectsBytes, rejectPage, 1.5)
      .then((url) => {
        setRejectDataUrl(url);
        setRejectLoadingPage(false);
      })
      .catch(() => {
        setRejectLoadingPage(false);
      });
  }, [rejectsBytes, rejectPage, rejectPageCount]);

  // ── File upload ───────────────────────────────────────────────────────────
  async function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const newPdfs: UploadedPdf[] = [];
    for (const file of files) {
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        const pages = await getPdfPageCount(bytes);
        newPdfs.push({
          id: `${Date.now()}-${Math.random()}`,
          filename: file.name,
          bytes,
          pages,
          processed: false,
        });
      } catch {
        addAlert("error", `Failed to read ${file.name}`);
      }
    }

    const updated = [...uploadedPdfs, ...newPdfs];
    setUploadedPdfs(updated);
    await saveProcessState({ uploadedPdfs: updated });
    await refreshStorage();
    if (fileInputRef.current) fileInputRef.current.value = "";
    addAlert("success", `Added ${newPdfs.length} file(s).`);
  }

  async function handleRemoveFile(id: string) {
    const updated = uploadedPdfs.filter((f) => f.id !== id);
    setUploadedPdfs(updated);
    await saveProcessState({ uploadedPdfs: updated });
  }

  async function handleClearUploads() {
    if (!window.confirm("Remove all uploaded files?")) return;
    setUploadedPdfs([]);
    await saveProcessState({ uploadedPdfs: [] });
  }

  async function handleToggleProcessed(id: string) {
    const updated = uploadedPdfs.map((f) =>
      f.id === id ? { ...f, processed: !f.processed } : f
    );
    setUploadedPdfs(updated);
    await saveProcessState({ uploadedPdfs: updated });
  }

  function handlePreviewUpload(pdf: UploadedPdf) {
    setModal({ kind: "pdf", bytes: pdf.bytes, title: pdf.filename });
  }

  // ── Process ───────────────────────────────────────────────────────────────
  async function handleProcess() {
    const toProcess = uploadedPdfs.filter((f) => !f.processed);
    if (toProcess.length === 0) {
      addAlert("warning", "No unprocessed files to scan.");
      return;
    }

    setProcessing(true);
    cancelRef.current = false;
    setProgress({ current: 0, total: 0, message: "Merging PDFs…" });

    try {
      // Merge all unprocessed PDFs into one
      const merged = await mergePdfs(toProcess.map((f) => f.bytes));

      const result = await sortScannedScoresheets(
        merged.buffer as ArrayBuffer,
        async (done, total) => {
          setProgress({ current: done, total, message: `Scanning page ${done} of ${total}…` });
        },
        () => cancelRef.current
      );

      if (cancelRef.current) {
        setProcessing(false);
        setProgress(null);
        addAlert("warning", "Processing stopped.");
        return;
      }

      const sorted: ProcessedPdf[] = result.sorted.map((s) => ({
        name: s.name,
        bytes: s.bytes,
      }));
      const rejects = result.rejects ?? null;

      // Merge with any previously processed PDFs (combine by judging number)
      const allProcessed = await mergeSortedPdfs(processedPdfs, sorted);
      setProcessedPdfs(allProcessed);

      // Mark uploaded files as processed
      const updatedUploads = uploadedPdfs.map((f) =>
        toProcess.find((tp) => tp.id === f.id) ? { ...f, processed: true } : f
      );
      setUploadedPdfs(updatedUploads);

      // Handle rejects: append to existing rejects if any
      let finalRejects: Uint8Array | null = rejects;
      if (rejectsBytes && rejects) {
        finalRejects = await mergePdfs([new Uint8Array(rejectsBytes), rejects]);
      }
      setRejectsBytes(finalRejects ?? null);
      if (finalRejects) {
        const count = await getPdfPageCount(finalRejects);
        setRejectPageCount(count);
        setRejectPage(0);
      } else {
        setRejectPageCount(0);
      }

      await saveProcessState({
        uploadedPdfs: updatedUploads,
        processedPdfs: allProcessed,
        rejectsBytes: finalRejects ?? null,
      });

      setProcessing(false);
      setProgress(null);
      await refreshStorage();
      addAlert(
        "success",
        `Processed ${result.processed} page(s), ${result.rejected} rejected.`
      );
    } catch (err) {
      setProcessing(false);
      setProgress(null);
      addAlert("error", `Processing failed: ${err}`);
    }
  }

  function handleStopProcessing() {
    cancelRef.current = true;
  }

  // ── Reject page navigation ────────────────────────────────────────────────
  async function handleDeleteRejectPage() {
    if (!rejectsBytes) return;
    if (!window.confirm("Delete this page permanently?")) return;
    const updated = await deletePage(rejectsBytes, rejectPage);
    if (!updated) {
      setRejectsBytes(null);
      setRejectPageCount(0);
      setRejectPage(0);
      await saveProcessState({ rejectsBytes: null });
    } else {
      const count = await getPdfPageCount(updated);
      setRejectsBytes(updated);
      setRejectPageCount(count);
      setRejectPage(Math.min(rejectPage, count - 1));
      await saveProcessState({ rejectsBytes: updated });
    }
    await refreshStorage();
  }

  async function handleAssignRejectPage() {
    if (!rejectsBytes || !assignInput.trim()) return;
    const judgingNumber = padJudgingNumber(assignInput);
    try {
      const existing = processedPdfs.find((p) => p.name === `${judgingNumber}.pdf`) ?? null;
      const { updatedRejects, updatedTarget } = await assignRejectPage(
        rejectsBytes,
        rejectPage,
        judgingNumber,
        existing?.bytes ?? null
      );

      // Update processed PDFs
      const allProcessed = processedPdfs.filter((p) => p.name !== `${judgingNumber}.pdf`);
      allProcessed.push({ name: `${judgingNumber}.pdf`, bytes: updatedTarget });
      setProcessedPdfs(allProcessed);

      setRejectsBytes(updatedRejects);
      if (!updatedRejects) {
        setRejectPageCount(0);
        setRejectPage(0);
      } else {
        const count = await getPdfPageCount(updatedRejects);
        setRejectPageCount(count);
        setRejectPage(Math.min(rejectPage, count - 1));
      }

      await saveProcessState({ rejectsBytes: updatedRejects, processedPdfs: allProcessed });
      await refreshStorage();
      setAssignInput("");
      addAlert("success", `Page assigned to judging number ${judgingNumber}.`);
    } catch (err) {
      addAlert("error", `Failed to assign page: ${err}`);
    }
  }

  // ── Processed file actions ────────────────────────────────────────────────
  function handleDownloadProcessed(pdf: ProcessedPdf) {
    const url = URL.createObjectURL(
      new Blob([new Uint8Array(pdf.bytes)], { type: "application/pdf" })
    );
    const a = document.createElement("a");
    a.href = url;
    a.download = pdf.name;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function handleDownloadAllProcessed() {
    if (processedPdfs.length === 0) return;
    addAlert("info", "Building ZIP archive…");
    try {
      const zip = new JSZip();
      for (const pdf of processedPdfs) {
        zip.file(pdf.name, pdf.bytes);
      }
      if (rejectsBytes) {
        zip.file("_rejects.pdf", rejectsBytes);
      }
      const blob = await zip.generateAsync({ type: "blob" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "sorted_scoresheets.zip";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      addAlert("error", `Failed to create ZIP: ${err}`);
    }
  }

  async function handleDeleteProcessed(name: string) {
    if (!window.confirm(`Delete ${name}?`)) return;
    const updated = processedPdfs.filter((p) => p.name !== name);
    setProcessedPdfs(updated);
    await saveProcessState({ processedPdfs: updated });
  }

  async function handleClearAllProcessed() {
    if (!window.confirm("Delete all sorted scoresheets and rejected pages?")) return;
    setProcessedPdfs([]);
    setRejectsBytes(null);
    setRejectPageCount(0);
    setRejectPage(0);
    await saveProcessState({ processedPdfs: [], rejectsBytes: null });
  }

  async function handleReset() {
    if (!window.confirm("Clear all session data?")) return;
    setUploadedPdfs([]);
    setProcessedPdfs([]);
    setRejectsBytes(null);
    setRejectPageCount(0);
    setRejectPage(0);
    setRejectDataUrl(null);
    await clearProcessState();
    await refreshStorage();
  }

  // ── Helpers ───────────────────────────────────────────────────────────────
  const hasUnprocessed = uploadedPdfs.some((f) => !f.processed);
  const totalUploadedPages = uploadedPdfs.reduce((sum, f) => sum + f.pages, 0);

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <>
      <AlertMessage alerts={alerts} onDismiss={dismissAlert} />

      {/* ── Upload Section ── */}
      <div className="card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h2>Upload Scanned Scoresheets</h2>
          {(uploadedPdfs.length > 0 || processedPdfs.length > 0) && (
            <button className="btn btn-small btn-danger" onClick={handleReset}>
              Reset Session
            </button>
          )}
        </div>
        {(uploadedPdfs.length > 0 || processedPdfs.length > 0) && storageUsage > 0 && (
          <p style={{ textAlign: "right", fontSize: "0.8rem", color: "var(--text-muted)", marginTop: 2, marginBottom: 0 }}>
            Session storage: {formatFileSize(storageUsage)}
          </p>
        )}
        <p className="form-hint">Select one or more PDF files. You can add more files before processing.</p>

        <div className="form-group">
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf"
            multiple
            onChange={handleFileChange}
          />
        </div>

        {uploadedPdfs.length > 0 ? (
          <div>
            <div className="alert alert-success" style={{ marginBottom: 12 }}>
              {uploadedPdfs.length} file(s) uploaded &mdash; {totalUploadedPages} total pages
            </div>
            <table className="upload-table">
              <thead>
                <tr>
                  <th>Status</th>
                  <th>Filename</th>
                  <th>Pages</th>
                  <th>Size</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {uploadedPdfs.map((f) => (
                  <tr key={f.id}>
                    <td>
                      <span
                        className={`status-icon ${f.processed ? "status-processed" : "status-pending"}`}
                        title={f.processed ? "Processed — click to re-enable" : "Pending"}
                        onClick={() => handleToggleProcessed(f.id)}
                        style={{ cursor: "pointer" }}
                      >
                        {f.processed ? "\u2713" : "\u25CB"}
                      </span>
                    </td>
                    <td>{f.filename}</td>
                    <td>{f.pages}</td>
                    <td>{formatFileSize(f.bytes.length)}</td>
                    <td className="file-actions">
                      <button
                        className="btn-icon"
                        title="Preview"
                        onClick={() => handlePreviewUpload(f)}
                      >
                        &#128269;
                      </button>
                      <button
                        className="btn-icon danger"
                        title="Remove"
                        onClick={() => handleRemoveFile(f.id)}
                      >
                        &times;
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <th></th>
                  <th>Total</th>
                  <th>{totalUploadedPages}</th>
                  <th>{formatFileSize(uploadedPdfs.reduce((s, f) => s + f.bytes.length, 0))}</th>
                  <th></th>
                </tr>
              </tfoot>
            </table>
            <div style={{ marginTop: 8 }}>
              <button className="btn btn-danger btn-small" onClick={handleClearUploads}>
                Clear All
              </button>
            </div>
          </div>
        ) : (
          <span className="file-status">No files uploaded</span>
        )}
      </div>

      {/* ── Process Section ── */}
      <div className="card">
        <h2>Process Scoresheets</h2>
        <p>Scan QR codes from uploaded scoresheets and sort them by judging number.</p>

        <div className="form-group" style={{ marginTop: 16 }}>
          <button
            className="btn btn-primary"
            onClick={handleProcess}
            disabled={!hasUnprocessed || processing}
          >
            {processing ? "Processing…" : "Process Scoresheets"}
          </button>
        </div>

        {processing && progress && (
          <ProgressBar
            current={progress.current}
            total={progress.total}
            message={progress.message}
            onStop={handleStopProcessing}
          />
        )}
      </div>

      {/* ── Rejects Viewer ── */}
      {rejectsBytes && rejectPageCount > 0 && (
        <div className="card">
          <h2>Review Rejected Scans</h2>
          <p>
            These pages could not be automatically processed. Review and manually assign judging
            numbers.
          </p>

          <div className="reject-viewer" style={{ marginTop: 16 }}>
            {/* Navigation */}
            <div className="reject-nav" style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
              <label>Page:</label>
              <select
                value={rejectPage}
                onChange={(e) => setRejectPage(Number(e.target.value))}
              >
                {Array.from({ length: rejectPageCount }, (_, i) => (
                  <option key={i} value={i}>
                    {i + 1} of {rejectPageCount}
                  </option>
                ))}
              </select>
              <span className="file-status">{rejectPageCount} page(s) need review</span>
            </div>

            {/* Page image */}
            <div style={{ textAlign: "center", margin: "20px 0" }}>
              {rejectLoadingPage ? (
                <div className="loading">Loading page…</div>
              ) : rejectDataUrl ? (
                <img
                  src={rejectDataUrl}
                  alt={`Rejected page ${rejectPage + 1}`}
                  className="reject-image preview-image clickable"
                  style={{ maxHeight: 500, cursor: "pointer" }}
                  onClick={() =>
                    setModal({
                      kind: "image",
                      src: rejectDataUrl,
                      title: `Rejected page ${rejectPage + 1} of ${rejectPageCount}`,
                    })
                  }
                />
              ) : null}
            </div>

            {/* Actions */}
            <div className="reject-actions" style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
              <button className="btn btn-danger" onClick={handleDeleteRejectPage}>
                Delete Page
              </button>
              <span style={{ color: "var(--text-muted)" }}>or assign judging number:</span>
              <input
                type="text"
                value={assignInput}
                onChange={(e) => setAssignInput(e.target.value.replace(/\D/g, "").slice(0, 6))}
                placeholder="e.g. 001234"
                maxLength={6}
                style={{ width: 100 }}
                onKeyDown={(e) => { if (e.key === "Enter") handleAssignRejectPage(); }}
              />
              <button
                className="btn btn-primary"
                onClick={handleAssignRejectPage}
                disabled={!assignInput.trim()}
              >
                Assign
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Sorted/Processed Files ── */}
      {processedPdfs.length > 0 && (
        <div className="card">
          <h2>Sorted Scoresheets</h2>
          <table className="upload-table">
            <thead>
              <tr>
                <th>Filename</th>
                <th>Size</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {[...processedPdfs]
                .sort((a, b) => a.name.localeCompare(b.name))
                .map((pdf) => (
                  <tr key={pdf.name}>
                    <td>{pdf.name}</td>
                    <td>{formatFileSize(pdf.bytes.length)}</td>
                    <td>
                      <span className="status-processed">&#10003; Processed</span>
                    </td>
                    <td className="file-actions">
                      <button
                        className="btn-icon"
                        title="Preview"
                        onClick={() => setModal({ kind: "pdf", bytes: pdf.bytes, title: pdf.name })}
                      >
                        &#128269;
                      </button>
                      <button
                        className="btn-icon download"
                        title="Download"
                        onClick={() => handleDownloadProcessed(pdf)}
                      >
                        &#11015;
                      </button>
                      <button
                        className="btn-icon danger"
                        title="Delete"
                        onClick={() => handleDeleteProcessed(pdf.name)}
                      >
                        &times;
                      </button>
                    </td>
                  </tr>
                ))}
            </tbody>
            <tfoot>
              <tr>
                <th>Total ({processedPdfs.length} files)</th>
                <th>{formatFileSize(processedPdfs.reduce((s, p) => s + p.bytes.length, 0))}</th>
                <th></th>
                <th></th>
              </tr>
            </tfoot>
          </table>
          <div style={{ display: "flex", gap: 12, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={handleDownloadAllProcessed}>
              Download All (ZIP)
            </button>
            <button className="btn btn-danger" onClick={handleClearAllProcessed}>
              Clear All
            </button>
          </div>
        </div>
      )}

      <Modal content={modal} onClose={() => setModal(null)} />
    </>
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────

/**
 * Merge two lists of ProcessedPdf, combining duplicates by appending pages.
 * This is done in memory (pdf-lib) — for the in-process use case this is fine.
 */
async function mergeSortedPdfs(
  existing: ProcessedPdf[],
  incoming: ProcessedPdf[]
): Promise<ProcessedPdf[]> {
  const map = new Map<string, Uint8Array>();
  for (const p of existing) map.set(p.name, p.bytes);

  for (const p of incoming) {
    if (map.has(p.name)) {
      // Append pages
      const merged = await mergePdfs([map.get(p.name)!, p.bytes]);
      map.set(p.name, merged);
    } else {
      map.set(p.name, p.bytes);
    }
  }

  return [...map.entries()].map(([name, bytes]) => ({ name, bytes }));
}
