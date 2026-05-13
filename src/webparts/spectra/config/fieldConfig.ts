// ─────────────────────────────────────────────────────────────────
// SPECTRA Document Repository — Field Configuration
//
// Defines per-document-type field requirements, visibility, and
// single vs. multi-select rules based on BRD metadata grid.
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
// Field Status Enum
// ─────────────────────────────────────────────────────────────────
export enum FieldRequirement {
  REQUIRED_SINGLE = "REQUIRED_SINGLE",
  REQUIRED_MULTI = "REQUIRED_MULTI",
  OPTIONAL_SINGLE = "OPTIONAL_SINGLE",
  OPTIONAL_MULTI = "OPTIONAL_MULTI",
  HIDDEN = "HIDDEN",
  SYSTEM_GENERATED = "SYSTEM_GENERATED",
}

// ─────────────────────────────────────────────────────────────────
// Field Names Constants
// ─────────────────────────────────────────────────────────────────
export const FIELD_NAMES = {
  DOCUMENT_TYPE: "documentType",
  THERAPEUTIC_AREA: "therapeuticArea",
  SUB_THERAPEUTIC_AREA: "subTherapeuticArea",
  DISEASE_AREA: "diseaseArea",
  INDICATION: "indication",
  LINE_OF_THERAPY: "lineOfTherapy",
  ASSET: "asset",
  PAID: "paid",
  EFFECTIVE_DATE: "effectiveDate",
  DESCRIPTION: "description",
  COMMENTS: "comments",
  STATUS: "status",
  UPLOAD_DATE: "uploadDate",
  ORG_FILE_NAME: "orgFileName",
  IMMUTABLE_FILE_NAME: "immutableFileName",
} as const;

// ─────────────────────────────────────────────────────────────────
// Field Configuration per Document Type (BRD Metadata Grid)
// Release 1.0 Document Types Only
// ─────────────────────────────────────────────────────────────────

type FieldConfigMap = Record<string, FieldRequirement>;

const TA_DRIVEN_SUB_TA_DOC_TYPES = new Set([
  "TPP",
  "TPC",
  "IEP",
  "EBP",
  "EIVP",
  "IAS",
]);

const TA_DRIVEN_LOT_DOC_TYPES = new Set([
  "TPP",
  "TPC",
  "IEP",
  "EBP",
  "EIVP",
  "IAS",
]);

const fieldConfigByDocumentType: Record<string, FieldConfigMap> = {
  // ─ DAS: Disease Area Strategy ─
  DAS: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.REQUIRED_MULTI,
    [FIELD_NAMES.INDICATION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.ASSET]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.PAID]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ TPP: Target Product Profile ─
  TPP: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.INDICATION]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_MULTI for ONC only
    [FIELD_NAMES.ASSET]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.PAID]: FieldRequirement.REQUIRED_MULTI,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ TPC: Target Product Claims ─
  TPC: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.INDICATION]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_MULTI for ONC only
    [FIELD_NAMES.ASSET]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.PAID]: FieldRequirement.REQUIRED_MULTI,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ IEP: Integrated Evidence Plan ─
  IEP: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.INDICATION]: FieldRequirement.REQUIRED_MULTI, // Different from TPP/TPC which are single
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_MULTI for ONC only
    [FIELD_NAMES.ASSET]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.PAID]: FieldRequirement.REQUIRED_MULTI,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ EBP: Early Brand Plan ─
  EBP: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.INDICATION]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_MULTI for ONC only
    [FIELD_NAMES.ASSET]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.PAID]: FieldRequirement.OPTIONAL_SINGLE, // Different from others: Optional instead of Required
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ EIVP: Early Integrated Value Proposition ─
  EIVP: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.INDICATION]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_MULTI for ONC only
    [FIELD_NAMES.ASSET]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.PAID]: FieldRequirement.REQUIRED_MULTI,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ IAS: Integrated Access Strategy ─
  IAS: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_SINGLE for AES and ONC
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.INDICATION]: FieldRequirement.REQUIRED_MULTI, // Multi-select like IEP
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN, // Conditional: REQUIRED_MULTI for ONC only
    [FIELD_NAMES.ASSET]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.PAID]: FieldRequirement.REQUIRED_MULTI,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ PDS: Pipeline Development Summary ─
  PDS: {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.INDICATION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.ASSET]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.PAID]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },

  // ─ R&D One-Pager ─
  "R&D One-Pager": {
    [FIELD_NAMES.DOCUMENT_TYPE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.THERAPEUTIC_AREA]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.SUB_THERAPEUTIC_AREA]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.DISEASE_AREA]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.INDICATION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.LINE_OF_THERAPY]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.ASSET]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.PAID]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.EFFECTIVE_DATE]: FieldRequirement.REQUIRED_SINGLE,
    [FIELD_NAMES.DESCRIPTION]: FieldRequirement.HIDDEN,
    [FIELD_NAMES.COMMENTS]: FieldRequirement.OPTIONAL_SINGLE,
    [FIELD_NAMES.STATUS]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.UPLOAD_DATE]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.ORG_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
    [FIELD_NAMES.IMMUTABLE_FILE_NAME]: FieldRequirement.SYSTEM_GENERATED,
  },
};

// ─────────────────────────────────────────────────────────────────
// Exports: Utility Functions
// ─────────────────────────────────────────────────────────────────

/**
 * Get field requirement status for a specific field in a document type.
 * Handles TA-driven visibility for Sub-TA (AES/ONC only) and LOT (ONC only).
 *
 * @param fieldName - Field name constant (e.g., FIELD_NAMES.INDICATION)
 * @param documentType - Document type (e.g., "TPP")
 * @param selectedTA - Selected therapeutic area (optional, for TA-dependent visibility)
 * @returns FieldRequirement enum value
 */
export const getFieldRequirement = (
  fieldName: string,
  documentType: string,
  selectedTA?: string,
): FieldRequirement => {
  const config = fieldConfigByDocumentType[documentType];
  if (!config) return FieldRequirement.HIDDEN;

  let requirement = config[fieldName];

  // TA-driven visibility: Sub-TA visible for Aesthetics and Oncology
  if (
    fieldName === FIELD_NAMES.SUB_THERAPEUTIC_AREA &&
    selectedTA &&
    TA_DRIVEN_SUB_TA_DOC_TYPES.has(documentType) &&
    ["aesthetics", "oncology"].includes(selectedTA.toLowerCase())
  ) {
    // If base config says HIDDEN, upgrade to REQUIRED_SINGLE for supported TAs
    if (requirement === FieldRequirement.HIDDEN) {
      requirement = FieldRequirement.REQUIRED_SINGLE;
    }
  } else if (fieldName === FIELD_NAMES.SUB_THERAPEUTIC_AREA && selectedTA) {
    // Other TAs: always hidden
    requirement = FieldRequirement.HIDDEN;
  }

  // TA-driven visibility: LOT only visible for Oncology
  if (
    fieldName === FIELD_NAMES.LINE_OF_THERAPY &&
    selectedTA &&
    TA_DRIVEN_LOT_DOC_TYPES.has(documentType) &&
    selectedTA.toLowerCase() === "oncology"
  ) {
    // If base config says HIDDEN, upgrade to REQUIRED_MULTI for Oncology
    if (requirement === FieldRequirement.HIDDEN) {
      requirement = FieldRequirement.REQUIRED_MULTI;
    } else if (requirement === FieldRequirement.REQUIRED_SINGLE) {
      requirement = FieldRequirement.REQUIRED_MULTI;
    } else if (requirement === FieldRequirement.OPTIONAL_SINGLE) {
      requirement = FieldRequirement.OPTIONAL_MULTI;
    }
  } else if (fieldName === FIELD_NAMES.LINE_OF_THERAPY && selectedTA) {
    // Non-Oncology TAs: always hidden
    requirement = FieldRequirement.HIDDEN;
  }

  return requirement;
};

/**
 * Check if a field should be visible in the form.
 * Hidden or System-Generated fields return false.
 *
 * @param fieldName - Field name constant
 * @param documentType - Document type
 * @param selectedTA - Selected therapeutic area (optional)
 * @returns true if field should be shown
 */
export const isFieldVisible = (
  fieldName: string,
  documentType: string,
  selectedTA?: string,
): boolean => {
  const requirement = getFieldRequirement(fieldName, documentType, selectedTA);
  return (
    requirement !== FieldRequirement.HIDDEN &&
    requirement !== FieldRequirement.SYSTEM_GENERATED
  );
};

/**
 * Check if a field is required (user must provide a value).
 * Required fields cannot be empty; optional/hidden/system fields can.
 *
 * @param fieldName - Field name constant
 * @param documentType - Document type
 * @param selectedTA - Selected therapeutic area (optional)
 * @returns true if field must be filled
 */
export const isFieldRequired = (
  fieldName: string,
  documentType: string,
  selectedTA?: string,
): boolean => {
  const requirement = getFieldRequirement(fieldName, documentType, selectedTA);
  return (
    requirement === FieldRequirement.REQUIRED_SINGLE ||
    requirement === FieldRequirement.REQUIRED_MULTI
  );
};

/**
 * Check if a field should be multi-select (returns multiple values).
 * Single-select fields return one value; multi-select return arrays.
 *
 * @param fieldName - Field name constant
 * @param documentType - Document type
 * @param selectedTA - Selected therapeutic area (optional)
 * @returns true if field supports multiple selections
 */
export const isFieldMultiSelect = (
  fieldName: string,
  documentType: string,
  selectedTA?: string,
): boolean => {
  const requirement = getFieldRequirement(fieldName, documentType, selectedTA);
  return (
    requirement === FieldRequirement.REQUIRED_MULTI ||
    requirement === FieldRequirement.OPTIONAL_MULTI
  );
};

/**
 * Get all field configuration for a document type.
 * Useful for bulk operations or UI rendering loops.
 *
 * @param documentType - Document type
 * @param selectedTA - Selected therapeutic area (optional)
 * @returns Record of fieldName → FieldRequirement
 */
export const getFieldConfig = (
  documentType: string,
  selectedTA?: string,
): Record<string, FieldRequirement> => {
  const config = fieldConfigByDocumentType[documentType];
  if (!config) return {};

  // If TA provided, recalculate TA-driven visibility
  if (!selectedTA) return config;

  const adjustedConfig = { ...config };
  adjustedConfig[FIELD_NAMES.SUB_THERAPEUTIC_AREA] = getFieldRequirement(
    FIELD_NAMES.SUB_THERAPEUTIC_AREA,
    documentType,
    selectedTA,
  );
  adjustedConfig[FIELD_NAMES.LINE_OF_THERAPY] = getFieldRequirement(
    FIELD_NAMES.LINE_OF_THERAPY,
    documentType,
    selectedTA,
  );

  return adjustedConfig;
};
