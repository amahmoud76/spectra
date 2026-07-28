import { useState, useEffect, useCallback, useMemo } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DocumentService } from "../services/DocumentService";

// ─────────────────────────────────────────────────────────────
// Shape of the aggregated summary data
// ─────────────────────────────────────────────────────────────

export interface IDocumentSummaryData {
  /** matrix[ta][docType] = count of Current documents */
  matrix: Record<string, Record<string, number>>;
  /** rowTotals[ta] = total across all document types for that TA */
  rowTotals: Record<string, number>;
  /** colTotals[docType] = total across all TAs for that document type */
  colTotals: Record<string, number>;
  /** Grand total (sum of all cells; multi-TA docs are counted per TA row) */
  grandTotal: number;
}

export interface IUseDocumentSummaryResult {
  summary: IDocumentSummaryData;
  isLoading: boolean;
  isError: boolean;
  refetch: () => void;
}

const EMPTY_SUMMARY: IDocumentSummaryData = {
  matrix: {},
  rowTotals: {},
  colTotals: {},
  grandTotal: 0,
};

// ─────────────────────────────────────────────────────────────
// Hook
// ─────────────────────────────────────────────────────────────

/**
 * Fetch and aggregate a live Document Summary (TA × DocumentType counts).
 *
 * - Counts Active/Current documents only (never archived).
 * - Results are NEVER cached — fresh data on every `refetch()` call.
 * - The fetch fires once on mount and again whenever `refetch()` is called
 *   (e.g. when the user navigates back to the landing page).
 * - Axis values (taValues, docTypeValues) come from the already-loaded
 *   metadata; when they change, the matrix is re-aggregated without a
 *   second network request.
 *
 * @param context        - SPFx web part context
 * @param documentLibrary - Document library name (defaults to config)
 * @param taValues        - Canonical Therapeutic Area values for row axis
 * @param docTypeValues   - Canonical Document Type values for column axis
 * @param useMock         - Use fixture data instead of live SharePoint
 */
export const useDocumentSummary = (
  context: WebPartContext,
  documentLibrary: string | undefined,
  taValues: string[],
  docTypeValues: string[],
  useMock: boolean,
): IUseDocumentSummaryResult => {
  const [rawRows, setRawRows] = useState<
    Array<{ documentType: string; therapeuticArea: string[] }>
  >([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isError, setIsError] = useState(false);
  const [fetchTrigger, setFetchTrigger] = useState(0);

  const refetch = useCallback(() => setFetchTrigger((n) => n + 1), []);

  // ── Fetch raw rows (light projection — only 2 fields) ───────
  // Only re-fetches when fetch-relevant params change; axis values
  // do NOT trigger a second network request.
  useEffect(() => {
    let cancelled = false;

    void (async () => {
      setIsLoading(true);
      setIsError(false);
      try {
        const svc = new DocumentService(context, documentLibrary, useMock);
        const rows = await svc.getDocumentSummaryRows();
        if (!cancelled) setRawRows(rows);
      } catch {
        if (!cancelled) setIsError(true);
      } finally {
        if (!cancelled) setIsLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [fetchTrigger, context, documentLibrary, useMock]);

  // ── Aggregate (pure derivation — no network) ────────────────
  // Re-runs when raw rows arrive OR when axis values load from metadata.
  const summary = useMemo<IDocumentSummaryData>(() => {
    if (taValues.length === 0 || docTypeValues.length === 0) {
      return EMPTY_SUMMARY;
    }

    const matrix: Record<string, Record<string, number>> = {};
    const rowTotals: Record<string, number> = {};
    const colTotals: Record<string, number> = {};
    let grandTotal = 0;

    // Initialise all cells to 0 so the table always renders the full grid
    for (const ta of taValues) {
      matrix[ta] = {};
      rowTotals[ta] = 0;
      for (const dt of docTypeValues) {
        matrix[ta][dt] = 0;
      }
    }
    for (const dt of docTypeValues) {
      colTotals[dt] = 0;
    }

    // Use Sets for O(1) membership checks
    const taSet = new Set(taValues);
    const dtSet = new Set(docTypeValues);

    for (const row of rawRows) {
      const dt = row.documentType;
      if (!dtSet.has(dt)) continue;
      for (const ta of row.therapeuticArea) {
        if (!taSet.has(ta)) continue;
        matrix[ta][dt]++;
        rowTotals[ta]++;
        colTotals[dt]++;
        grandTotal++;
      }
    }

    return { matrix, rowTotals, colTotals, grandTotal };
  }, [rawRows, taValues, docTypeValues]);

  return { summary, isLoading, isError, refetch };
};
