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
export const generateFileName = (
  payload: Partial<IUploadPayload>,
  extension: string,
): string => {
  const ext = extension || "pdf";

  // Document type abbreviation
  const docTypeName = payload.documentType ?? "";
  const docType =
    DOC_TYPE_ABBREVIATIONS[docTypeName] || sanitize(docTypeName);

  // Therapeutic area abbreviation
  const taValues = payload.therapeuticArea ?? [];
  const ta =
    taValues.length > 0
      ? taValues
          .map((t) => TA_ABBREVIATIONS[t] || sanitize(t))
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Sub-TA for therapeutic areas that require it
  const includeSubTA = taValues.some((t) => TA_INCLUDES_SUB_TA.includes(t));
  const subTAValues = payload.subTherapeuticArea ?? [];
  const subTA =
    includeSubTA && subTAValues.length > 0
      ? subTAValues
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  const diseaseAreaValues = payload.diseaseArea ?? [];
  const diseaseArea =
    diseaseAreaValues.length > 0
      ? diseaseAreaValues
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  if (docTypeName === "DAS") {
    const segments: string[] = [docType];
    if (ta) segments.push(ta);
    if (subTA) segments.push(subTA);
    if (diseaseArea) segments.push(diseaseArea);
    return `${segments.join("-")}.${ext}`;
  }

  // Indication
  const indicationValues = payload.indication ?? [];
  const indication =
    indicationValues.length > 0
      ? indicationValues
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Line of Therapy (included if it exists)
  const lotValues = payload.lineOfTherapy ?? [];
  const lot =
    lotValues.length > 0
      ? lotValues
          .map(sanitize)
          .sort()
          .join(MULTI_VALUE_FILENAME_DELIMITER)
      : "";

  // Asset
  const assetValues = payload.asset ?? [];
  const asset =
    assetValues.length > 0
      ? assetValues
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
