import * as React from "react";
import styles from "../SPECTRA.module.scss";
import { IDocumentSummaryData } from "../../hooks/useDocumentSummary";

export interface IDocumentSummaryTableProps {
  taValues: string[];
  docTypeValues: string[];
  summary: IDocumentSummaryData;
  isLoading: boolean;
  isError: boolean;
  onRetry: () => void;
}

export const DocumentSummaryTable: React.FC<IDocumentSummaryTableProps> = ({
  taValues,
  docTypeValues,
  summary,
  isLoading,
  isError,
  onRetry,
}) => {
  if (isLoading) {
    return (
      <div className={styles.docSummarySection}>
        <div className={styles.docSummaryLoading}>
          <span className={styles.spinner} role="status" aria-label="Loading" />
          <span>Loading document summary&hellip;</span>
        </div>
      </div>
    );
  }

  if (isError) {
    return (
      <div className={styles.docSummarySection}>
        <div className={styles.docSummaryError}>
          <span>Unable to load document summary.</span>
          <button
            type="button"
            className={styles.docSummaryRetryBtn}
            onClick={onRetry}
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  const { matrix, rowTotals, colTotals, grandTotal } = summary;

  return (
    <div className={styles.docSummarySection}>
      <div className={styles.docSummaryScroll}>
        <table className={styles.docSummaryTable}>
          <thead>
            <tr>
              {/* Top-left: "DOCUMENT SUMMARY" spanning the two TA-axis header cells */}
              <th
                rowSpan={2}
                colSpan={2}
                className={styles.docSummarySectionLabel}
              >
                Document Summary
              </th>
              {/* Top-right: "Document Types" spanning all doc-type columns + Total */}
              <th
                colSpan={docTypeValues.length + 1}
                className={styles.docSummaryColGroupHeader}
              >
                Document Types
              </th>
            </tr>
            <tr>
              {docTypeValues.map((dt) => (
                <th key={dt} className={styles.docSummaryColHeader}>
                  {dt}
                </th>
              ))}
              <th className={styles.docSummaryColHeader}>Total</th>
            </tr>
          </thead>
          <tbody>
            {taValues.map((ta, idx) => (
              <tr key={ta}>
                {/* Rotated "Therapeutic Area" axis label — only rendered once, spans all TA rows */}
                {idx === 0 && (
                  <th
                    rowSpan={taValues.length}
                    scope="rowgroup"
                    className={styles.docSummaryTaAxisCell}
                  >
                    Therapeutic Area
                  </th>
                )}
                <th scope="row" className={styles.docSummaryTaNameCell}>
                  {ta}
                </th>
                {docTypeValues.map((dt) => (
                  <td key={dt}>{matrix[ta]?.[dt] ?? 0}</td>
                ))}
                <td className={styles.docSummaryTotalCell}>
                  {rowTotals[ta] ?? 0}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr>
              <td colSpan={2} className={styles.docSummaryTotalLabel}>
                Total Documents
              </td>
              {docTypeValues.map((dt) => (
                <td key={dt} className={styles.docSummaryTotalCell}>
                  {colTotals[dt] ?? 0}
                </td>
              ))}
              <td className={styles.docSummaryTotalCell}>{grandTotal}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
};

export default DocumentSummaryTable;
