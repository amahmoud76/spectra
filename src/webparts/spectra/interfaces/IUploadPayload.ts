export interface IUploadPayload {
  // ── File ──────────────────────────────────────────────────────
  file: File; // The actual file object from the file picker

  // ── Required metadata ─────────────────────────────────────────
  documentType: string; // Single-select — required
  therapeuticArea: string[]; // Multi-value — required
  diseaseArea: string[]; // Multi-value — required for DAS

  // ── Optional metadata ─────────────────────────────────────────
  subTherapeuticArea: string[]; // Multi-value — shown for Aesthetics
  asset: string[]; // Multi-value
  indication: string[]; // Multi-value
  lineOfTherapy: string[]; // Multi-value — primarily Oncology
  paid: string[]; // Multi-value
  effectiveDate: string; // ISO date string
  description: string; // Document description
  comments: string; // Free text — max 100 characters

  // ── System-generated (computed before submit) ─────────────────
  searchTokens: string[]; // Synonym tokens built from asset, TA, indication
  // Computed by buildDocumentSearchTokens() in the
  // upload/replace panel before submission
}

export interface IUploadResult {
  success: boolean;
  documentId?: string; // SharePoint list item ID of the new document
  generatedFileName?: string; // Auto-generated filename for success message
  message: string; // Success or error message
  isDuplicateIdentity?: boolean; // True if an identical Current document already exists
  existingDocumentId?: string; // ID of the existing Current document (if isDuplicateIdentity is true)
  existingFileName?: string; // Filename of existing document for display
  isFileNameConflict?: boolean; // True when a file with the same generated name already exists
  conflictFileName?: string; // The generated file name that conflicted
}
