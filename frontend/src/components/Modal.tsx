import { useEffect, useRef, useState } from "react";
import type { ModalContent } from "../lib/types";
import { getPdfPageCount, renderPdfPageToDataUrl } from "../lib/pdf-processor";

interface Props {
  content: ModalContent | null;
  onClose: () => void;
}

export default function Modal({ content, onClose }: Props) {
  const backdropRef = useRef<HTMLDivElement>(null);

  // Close on Escape
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, [onClose]);

  // Prevent body scroll when open
  useEffect(() => {
    if (content) {
      document.body.style.overflow = "hidden";
    } else {
      document.body.style.overflow = "";
    }
    return () => {
      document.body.style.overflow = "";
    };
  }, [content]);

  if (!content) return null;

  const title =
    content.kind === "image" ? content.title : content.kind === "pdf" ? content.title : content.title;

  return (
    <div
      ref={backdropRef}
      className="modal open"
      onClick={(e) => {
        if (e.target === backdropRef.current) onClose();
      }}
    >
      <div className="modal-content">
        <button className="modal-close" onClick={onClose}>
          &times;
        </button>
        <div className="modal-header">{title}</div>
        <div className="modal-body">
          {content.kind === "image" && (
            <img src={content.src} alt={content.title} style={{ display: "block", maxWidth: "100%" }} />
          )}
          {content.kind === "pdf" && (
            <PdfViewer bytes={content.bytes} />
          )}
          {content.kind === "loading" && (
            <div className="modal-loading">
              <div className="spinner" />
              <p>Loading preview…</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Multi-page PDF viewer ─────────────────────────────────────────────────

function PdfViewer({ bytes }: { bytes: Uint8Array }) {
  const [pages, setPages] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [total, setTotal] = useState(0);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setPages([]);
    setTotal(0);

    (async () => {
      // Use pdf-lib for page count — no pdfjs worker needed
      const numPages = await getPdfPageCount(bytes);
      if (cancelled) return;
      setTotal(numPages);

      for (let i = 0; i < numPages; i++) {
        if (cancelled) break;
        // renderPdfPageToDataUrl handles worker initialisation internally
        const dataUrl = await renderPdfPageToDataUrl(bytes, i, 1.5);
        if (!cancelled) setPages((prev) => [...prev, dataUrl]);
      }

      if (!cancelled) setLoading(false);
    })();

    return () => {
      cancelled = true;
    };
  }, [bytes]);

  return (
    <div style={{ overflowY: "auto", maxHeight: "calc(95vh - 60px)" }}>
      {loading && pages.length === 0 && (
        <div className="modal-loading">
          <div className="spinner" />
          <p>Rendering pages…</p>
        </div>
      )}
      {pages.map((src, i) => (
        <div key={i} className="modal-page">
          <div className="modal-page-label">
            Page {i + 1} of {total || "?"}
          </div>
          <img src={src} alt={`Page ${i + 1}`} style={{ display: "block", maxWidth: "100%", margin: "0 auto" }} />
        </div>
      ))}
    </div>
  );
}
