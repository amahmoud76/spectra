import { IDocument } from "../interfaces/IDocument";
import { parseISO, format, isValid } from "date-fns";

/**
 * Export documents to CSV format (BRD 2.13, Appendix D).
 * Fields: Document Type, TA, Asset, Indication, PAID, Disease Area,
 * Comments, Effective Date, Document Description, Status, Upload Date,
 * Original File Name, Immutable File Name, Uploaded By (511)
 *
 * Uses date-fns for consistent date formatting in the CSV output.
 */
export const exportDocumentsToCSV = (documents: IDocument[]): void => {
  const headers = [
    "Document Name",
    "Document Type",
    "Therapeutic Area",
    "Sub-Therapeutic Area",
    "Asset",
    "Indication",
    "Line of Therapy",
    "PAID",
    "Disease Area",
    "Comments",
    "Effective Date",
    "Document Description",
    "Status",
    "Upload Date",
    "Original File Name",
    "Immutable File Name",
    "Uploaded By (511)",
    "Created By",
    "Modified By",
  ];

  const rows = documents.map((doc) => [
    escapeCSV(doc.fileName),
    escapeCSV(doc.documentType),
    escapeCSV(doc.therapeuticArea.join("; ")),
    escapeCSV(doc.subTherapeuticArea.join("; ")),
    escapeCSV(doc.asset.join("; ")),
    escapeCSV(doc.indication.join("; ")),
    escapeCSV(doc.lineOfTherapy.join("; ")),
    escapeCSV(doc.paid.join("; ")),
    escapeCSV(doc.diseaseArea.join("; ")),
    escapeCSV(doc.comments),
    escapeCSV(formatDateForCSV(doc.effectiveDate)),
    escapeCSV(doc.description),
    escapeCSV(doc.status),
    escapeCSV(formatDateForCSV(doc.uploadDate)),
    escapeCSV(doc.immutableFileName),
    escapeCSV(doc.fileName),
    escapeCSV(doc.spectra511),
    escapeCSV(doc.createdBy),
    escapeCSV(doc.modifiedBy),
  ]);

  const csvContent = [
    headers.join(","),
    ...rows.map((row) => row.join(",")),
  ].join("\n");

  const blob = new Blob(["\uFEFF" + csvContent], {
    type: "text/csv;charset=utf-8;",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `SPECTRA_Export_${format(new Date(), "yyyy-MM-dd")}.csv`;
  link.click();
  URL.revokeObjectURL(url);
};

/**
 * Format a date string for CSV output.
 * Returns consistent yyyy-MM-dd format regardless of input.
 */
function formatDateForCSV(iso: string): string {
  if (!iso) return "";
  const date = parseISO(iso);
  if (!isValid(date)) return iso;
  return format(date, "yyyy-MM-dd");
}

/**
 * Escape a value for CSV — wrap in quotes if it contains commas,
 * quotes, or newlines.
 */
function escapeCSV(value: string): string {
  if (!value) return '""';
  if (value.includes(",") || value.includes('"') || value.includes("\n")) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return `"${value}"`;
}
