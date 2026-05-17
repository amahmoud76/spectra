import { IUploadPayload } from "../interfaces/IUploadPayload";
import {
  DOC_TYPE_ABBREVIATIONS,
  TA_ABBREVIATIONS,
  TA_INCLUDES_SUB_TA,
} from "../config/config";

const MULTI_VALUE_FILENAME_DELIMITER = "__";

/**
 * Generate a standardized filename from metadata tags (BRD Appendix H).
 *
 * Format (non-DAS):
 *   DocType-TA-Asset-Indication[-SubTA][-LineOfTherapy].ext
 *
 * Format (DAS):
 *   DAS-TA[-SubTA]-DiseaseArea.ext
 *   Disease Area is required for DAS and included in the filename.
 *   Optional/hidden fields (Asset, Indication, LOT, PAID) are excluded.
 *
 * Rules:
 *   - Multiple values concatenated with '__' in alphabetical order
 *   - Asset is placed before Indication (non-DAS only)
 *   - Sub-TA included for therapeutic areas listed in TA_INCLUDES_SUB_TA
 *   - Line of Therapy included if it exists (non-DAS only)
 *   - Disease Area included for DAS only; optional on other types and excluded
 *   - Spaces replaced with hyphens in each segment
 */
export const generateFileName = (payload: IUploadPayload): string => {
  const ext = payload.file.name.split(".").pop() || "pdf";

  // Document type abbreviation
  const docType =
    DOC_TYPE_ABBREVIATIONS[payload.documentType] ||
    sanitize(payload.documentType);

  // Therapeutic area abbreviation
  const ta =
    payload.therapeuticArea.length > 0
      ? payload.therapeuticArea
          .map((t) => TA_ABBREVIATIONS[t] || sanitize(t))
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Sub-TA for therapeutic areas that require it
  const includeSubTA = payload.therapeuticArea.some((t) =>
    TA_INCLUDES_SUB_TA.includes(t),
  );
  const subTA =
    includeSubTA && payload.subTherapeuticArea.length > 0
      ? payload.subTherapeuticArea
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  const diseaseArea =
    payload.diseaseArea.length > 0
      ? payload.diseaseArea
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  if (payload.documentType === "DAS") {
    const segments: string[] = [docType];
    if (ta) segments.push(ta);
    if (subTA) segments.push(subTA);
    if (diseaseArea) segments.push(diseaseArea);
    return `${segments.join("-")}.${ext}`;
  }

  // Indication
  const indication =
    payload.indication.length > 0
      ? payload.indication
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Line of Therapy (included if it exists)
  const lot =
    payload.lineOfTherapy.length > 0
      ? payload.lineOfTherapy
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Asset
  const asset =
    payload.asset.length > 0
      ? payload.asset
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Build segments in order
  // Note: Effective Date is excluded from filename per Release 1 requirements.
  // Versioning is based on metadata identity only (excluding Effective Date and Comments).
  const segments: string[] = [docType];
  if (ta) segments.push(ta);
  if (asset) segments.push(asset);
  if (indication) segments.push(indication);
  if (subTA) segments.push(subTA);
  if (lot) segments.push(lot);

  return `${segments.join("-")}.${ext}`;
};

/**
 * Replace spaces with hyphens and remove special characters.
 */
const sanitize = (value: string): string =>
  value
    .trim()
    .replace(/\s+/g, "-")
    .replace(/[^a-zA-Z0-9\-+]/g, "");
