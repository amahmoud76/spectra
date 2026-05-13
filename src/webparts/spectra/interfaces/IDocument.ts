// ─────────────────────────────────────────────────────────────────
// IDocument — Maps to the SPECTRA Document content type (18 columns)
// SharePoint natively tracks Created By / Modified By — no custom
// user-attribution columns needed.
// ─────────────────────────────────────────────────────────────────

export interface IDocument {
  // ── SharePoint system fields ──────────────────────────────────
  id: string; // SharePoint list item ID
  fileName: string; // Name column — current filename (auto-renamed)
  fileExtension: string; // File extension (pdf, docx, pptx, xlsx)
  fileUrl: string; // Server-relative URL to the file
  fileSize: number; // File size in bytes
  createdBy: string; // SharePoint native Created By (display name)
  createdByEmail: string; // SharePoint native Created By (email/UPN)
  modifiedBy: string; // SharePoint native Modified By (display name)
  modifiedByEmail: string; // SharePoint native Modified By (email/UPN)
  uploadDate: string; // SharePoint Created date — ISO 8601

  // ── Business metadata (Appendix B) ────────────────────────────
  title: string; // Display title
  asset: string[]; // SPECTRAAsset — multi-value, parsed from semicolon-separated
  documentType: string; // SPECTRADocumentType — single-select choice
  therapeuticArea: string[]; // SPECTRATherapeuticArea — multi-value
  subTherapeuticArea: string[]; // SPECTRASubTherapeuticArea — multi-value
  indication: string[]; // SPECTRAIndication — multi-value
  lineOfTherapy: string[]; // SPECTRALineOfTherapy — multi-value
  paid: string[]; // SPECTRAPAID — multi-value
  diseaseArea: string[]; // SPECTRADiseaseArea — multi-value
  effectiveDate: string; // SPECTRAEffectiveDate — ISO date string
  status: "Current" | "Archive"; // SPECTRAStatus — document lifecycle status
  description: string; // SPECTRADescription — document description
  comments: string; // SPECTRAComments — free text, max 100 chars
  spectra511: string; // SPECTRAFiveEleven — uploader identifier (Release 1 email)

  // ── System-generated fields ───────────────────────────────────
  searchTokens: string[]; // SPECTRASearchTokens — comma-separated synonyms, parsed to array
  immutableFileName: string; // SPECTRAImmutableFileName — original filename, never changed
}
