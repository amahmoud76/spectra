import { IDocument } from "../interfaces/IDocument";
import { ISortState } from "../interfaces/ISortState";
import { parseISO } from "date-fns";

/**
 * Sort documents by a single-value column.
 * Multi-value columns (asset, therapeuticArea, etc.) are NOT sortable.
 * Returns a new sorted array — does not mutate the input.
 *
 * Uses date-fns parseISO for timezone-safe date sorting.
 */
export const sortDocuments = (
  documents: IDocument[],
  sortState: ISortState | null,
): IDocument[] => {
  if (!sortState) return documents;

  const { field, direction } = sortState;
  const multiplier = direction === "asc" ? 1 : -1;

  return [...documents].sort((a, b) => {
    let aVal: string | number = "";
    let bVal: string | number = "";

    switch (field) {
      case "fileName":
        aVal = a.fileName.toLowerCase();
        bVal = b.fileName.toLowerCase();
        break;
      case "comments":
        aVal = a.comments.toLowerCase();
        bVal = b.comments.toLowerCase();
        break;
      case "documentType":
        aVal = a.documentType.toLowerCase();
        bVal = b.documentType.toLowerCase();
        break;
      case "effectiveDate":
        aVal = parseISO(a.effectiveDate).getTime() || 0;
        bVal = parseISO(b.effectiveDate).getTime() || 0;
        break;
      case "uploadDate":
        aVal = parseISO(a.uploadDate).getTime() || 0;
        bVal = parseISO(b.uploadDate).getTime() || 0;
        break;
      case "createdBy":
        aVal = a.createdBy.toLowerCase();
        bVal = b.createdBy.toLowerCase();
        break;
      case "modifiedBy":
        aVal = a.modifiedBy.toLowerCase();
        bVal = b.modifiedBy.toLowerCase();
        break;
      case "lineOfTherapy":
        aVal = a.lineOfTherapy.join("; ").toLowerCase();
        bVal = b.lineOfTherapy.join("; ").toLowerCase();
        break;
      case "diseaseArea":
        aVal = a.diseaseArea.join("; ").toLowerCase();
        bVal = b.diseaseArea.join("; ").toLowerCase();
        break;
      case "status":
        aVal = a.status.toLowerCase();
        bVal = b.status.toLowerCase();
        break;
      default:
        return 0;
    }

    if (aVal < bVal) return -1 * multiplier;
    if (aVal > bVal) return 1 * multiplier;
    return 0;
  });
};
