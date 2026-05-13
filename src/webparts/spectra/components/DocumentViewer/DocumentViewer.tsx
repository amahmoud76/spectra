import * as React from "react";
import styles from "../SPECTRA.module.scss";

// ── Supported file categories ────────────────────────────────────
const OFFICE_EXTENSIONS = ["docx", "doc", "xlsx", "xls", "pptx", "ppt"];
const PDF_EXTENSION = "pdf";
const SUPPORTED_EXTENSIONS = [...OFFICE_EXTENSIONS, PDF_EXTENSION];

// Normalize file URLs from SharePoint items.
// In this app, FileRef is usually server-relative (/sites/.../file.docx).
const toAbsoluteFileUrl = (fileUrl: string, siteUrl: string): string => {
  if (/^https?:\/\//i.test(fileUrl)) return fileUrl;
  if (fileUrl.startsWith("/")) return `${window.location.origin}${fileUrl}`;
  return `${siteUrl.replace(/\/$/, "")}/${fileUrl.replace(/^\//, "")}`;
};

const toServerRelativeFileUrl = (fileUrl: string, siteUrl: string): string => {
  if (fileUrl.startsWith("/")) return fileUrl;
  if (/^https?:\/\//i.test(fileUrl)) {
    try {
      return new URL(fileUrl).pathname;
    } catch {
      return fileUrl;
    }
  }
  try {
    return new URL(fileUrl, `${siteUrl.replace(/\/$/, "")}/`).pathname;
  } catch {
    return fileUrl;
  }
};

interface IPdfLib {
  GlobalWorkerOptions: { workerSrc: string };
  getDocument: (params: { url: string; withCredentials: boolean }) => {
    promise: Promise<IPdfDocument>;
  };
}

interface IPdfDocument {
  numPages: number;
  getPage: (n: number) => Promise<IPdfPage>;
}

interface IPdfPage {
  getViewport: (opts: { scale: number }) => { width: number; height: number };
  render: (opts: {
    canvasContext: CanvasRenderingContext2D;
    viewport: { width: number; height: number };
  }) => { promise: Promise<void>; cancel: () => void };
}

interface IPdfState {
  totalPages: number;
  currentPage: number;
  isLoading: boolean;
  error: string | null;
}

const PdfViewer: React.FC<{ fileUrl: string; siteUrl: string }> = ({
  fileUrl,
  siteUrl,
}) => {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const pdfDocRef = React.useRef<IPdfDocument | null>(null);
  const renderTaskRef = React.useRef<{ cancel: () => void } | null>(null);

  const [state, setState] = React.useState<IPdfState>({
    totalPages: 0,
    currentPage: 1,
    isLoading: true,
    error: null,
  });

  const absoluteUrl = toAbsoluteFileUrl(fileUrl, siteUrl);

  React.useEffect(() => {
    let cancelled = false;

    const load = async (): Promise<void> => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const pdfjsLib = require("pdfjs-dist/build/pdf") as IPdfLib;
        pdfjsLib.GlobalWorkerOptions.workerSrc = `${siteUrl}/SiteAssets/pdf.worker.js`;

        const loadingTask = pdfjsLib.getDocument({
          url: absoluteUrl,
          withCredentials: true,
        });

        const pdf = await loadingTask.promise;
        if (cancelled) return;

        pdfDocRef.current = pdf;
        setState((prev) => ({
          ...prev,
          totalPages: pdf.numPages,
          isLoading: false,
        }));
      } catch {
        if (!cancelled) {
          setState((prev) => ({
            ...prev,
            isLoading: false,
            error: "Unable to load PDF preview.",
          }));
        }
      }
    };

    void load();
    return () => {
      cancelled = true;
    };
  }, [absoluteUrl, siteUrl]);

  React.useEffect(() => {
    if (!pdfDocRef.current || !canvasRef.current) return;

    if (renderTaskRef.current) {
      renderTaskRef.current.cancel();
      renderTaskRef.current = null;
    }

    const render = async (): Promise<void> => {
      try {
        const pdf = pdfDocRef.current;
        if (!pdf) return;

        const page = await pdf.getPage(state.currentPage);
        const canvas = canvasRef.current;
        if (!canvas) return;

        const containerWidth = containerRef.current?.clientWidth || 900;
        const rawViewport = page.getViewport({ scale: 1 });
        const scale = containerWidth / rawViewport.width;
        const viewport = page.getViewport({ scale });

        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const ctx = canvas.getContext("2d");
        if (!ctx) return;

        const renderTask = page.render({ canvasContext: ctx, viewport });
        renderTaskRef.current = renderTask;
        await renderTask.promise;
        renderTaskRef.current = null;
      } catch {
        // Ignore cancelled render task errors
      }
    };

    void render();
  }, [state.currentPage, state.totalPages]);

  const goTo = (page: number): void => {
    if (page < 1 || page > state.totalPages) return;
    setState((prev) => ({ ...prev, currentPage: page }));
  };

  if (state.isLoading) {
    return (
      <div className={styles.viewerContainer}>
        <div className={styles.shimmer} style={{ height: 700 }} />
      </div>
    );
  }

  if (state.error) {
    return (
      <div className={styles.viewerUnsupported}>
        <i
          className={`fa fa-triangle-exclamation ${styles.viewerUnsupportedIcon}`}
        />
        <p>{state.error}</p>
      </div>
    );
  }

  return (
    <div className={styles.viewerContainer} ref={containerRef}>
      {state.totalPages > 1 && (
        <div className={styles.viewerPageControls}>
          <button
            className={styles.viewerPageBtn}
            onClick={() => goTo(1)}
            disabled={state.currentPage === 1}
            aria-label="First page"
            type="button"
          >
            First
          </button>
          <button
            className={styles.viewerPageBtn}
            onClick={() => goTo(state.currentPage - 1)}
            disabled={state.currentPage === 1}
            aria-label="Previous page"
            type="button"
          >
            Previous
          </button>
          <span className={styles.viewerPageInfo}>
            Page {state.currentPage} of {state.totalPages}
          </span>
          <button
            className={styles.viewerPageBtn}
            onClick={() => goTo(state.currentPage + 1)}
            disabled={state.currentPage === state.totalPages}
            aria-label="Next page"
            type="button"
          >
            Next
          </button>
          <button
            className={styles.viewerPageBtn}
            onClick={() => goTo(state.totalPages)}
            disabled={state.currentPage === state.totalPages}
            aria-label="Last page"
            type="button"
          >
            Last
          </button>
        </div>
      )}

      <canvas
        ref={canvasRef}
        className={styles.viewerCanvas}
        aria-label="PDF document preview"
      />
    </div>
  );
};

const OfficeViewer: React.FC<{ fileUrl: string; siteUrl: string }> = ({
  fileUrl,
  siteUrl,
}) => {
  const iframeRef = React.useRef<HTMLIFrameElement>(null);
  const sourceDoc = toServerRelativeFileUrl(fileUrl, siteUrl);

  // embedview should receive a server-relative file path in sourcedoc.
  // Query flags reduce visible Office actions; onLoad CSS injection hides remaining status/menu controls.
  const embedUrl = `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(sourceDoc)}&action=embedview&wdHideHeaders=true&wdDownloadButton=false&wdPrint=false`;

  const handleLoad = (): void => {
    try {
      const iframeDoc = iframeRef.current?.contentDocument;
      if (!iframeDoc || !iframeDoc.head) return;

      // Prevent duplicate style tags if iframe reloads.
      const existingStyle = iframeDoc.getElementById("spectra-wopi-lockdown-style");
      if (existingStyle) return;

      const style = iframeDoc.createElement("style");
      style.id = "spectra-wopi-lockdown-style";
      style.textContent = `
        #WACStatusBarContainer { display: none !important; }
        #ControlMenu-Small14 { display: none !important; }
        #fseaFullScreen-Small14 { display: none !important; }
        #WACAppBase { bottom: 0 !important; }
      `;
      iframeDoc.head.appendChild(style);
    } catch {
      // If browser/tenant isolation blocks iframe DOM access, keep fallback URL params only.
    }
  };

  return (
    <iframe
      ref={iframeRef}
      className={styles.pdfViewer}
      src={embedUrl}
      title="Document preview"
      frameBorder="0"
      sandbox="allow-scripts allow-same-origin allow-forms"
      onLoad={handleLoad}
      aria-label="Document preview"
    />
  );
};

export interface IDocumentViewerProps {
  fileName: string;
  fileUrl: string;
  siteUrl: string;
}

export const DocumentViewer: React.FC<IDocumentViewerProps> = ({
  fileName,
  fileUrl,
  siteUrl,
}) => {
  const ext = fileName.split(".").pop()?.toLowerCase() ?? "";

  if (!SUPPORTED_EXTENSIONS.includes(ext)) {
    return (
      <div className={styles.viewerUnsupported}>
        <i className={`fa fa-file ${styles.viewerUnsupportedIcon}`} />
        <p>Preview is not available for this file type.</p>
        <p className={styles.viewerUnsupportedSub}>
          Supported types: PDF, Word, Excel, PowerPoint
        </p>
      </div>
    );
  }

  if (ext === PDF_EXTENSION) {
    return <PdfViewer fileUrl={fileUrl} siteUrl={siteUrl} />;
  }

  return <OfficeViewer fileUrl={fileUrl} siteUrl={siteUrl} />;
};
