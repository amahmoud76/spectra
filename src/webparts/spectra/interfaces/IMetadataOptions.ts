// ─────────────────────────────────────────────────────────────────
// Single metadata option — used for dropdown items
// SearchTokens enable synonym search (e.g., searching "ibrutinib"
// finds documents tagged with asset "Imbruvica")
// ─────────────────────────────────────────────────────────────────
export interface IMetadataOption {
  value: string; // Display value (e.g., "Imbruvica")
  searchTokens: string[]; // Synonyms (e.g., ["ibrutinib", "btk inhibitor"])
}

// ─────────────────────────────────────────────────────────────────
// Cascading relationship — one row from SPECTRA_ProjectPAID list
// Used for cascading metadata filters (BRD 3.4):
//   Select PAID → auto-populate TA, Asset, Indication
//   Select TA → filter to relevant PAID, Asset, Indication
//   Select Asset → filter to relevant TA, PAID, Indication
//   Select Indication → filter to relevant TA, Asset, PAID
// ─────────────────────────────────────────────────────────────────
export interface IProjectPaidRelationship {
  projectPaid: string; // PROJECT_PAID value
  therapeuticArea: string; // THERAPEUTIC_AREA
  assetNumber: string; // ASSET_NUMBER
  indication: string; // INDICATION
  subTherapeuticArea?: string; // SUB_TA (optional, auto-populated from relationship)
  lineOfTherapy?: string; // LINE_OF_THERAPY (optional, auto-populated from relationship)
}

// ─────────────────────────────────────────────────────────────────
// Disease Area Strategy relationship — one row from SPECTRA_DAS list
// Used for TA -> DAS cascading in upload/replace/filter flows
// ─────────────────────────────────────────────────────────────────
export interface IDiseaseAreaStrategyRelationship {
  value: string; // DAS Title value
  therapeuticArea: string; // THERAPEUTIC_AREA
}

// ─────────────────────────────────────────────────────────────────
// All metadata options — loaded by MetadataService, cached 24hrs
// ─────────────────────────────────────────────────────────────────
export interface IMetadataOptions {
  assets: IMetadataOption[]; // From SPECTRA_Assets list
  documentTypes: IMetadataOption[]; // From SPECTRA_DocumentTypes list
  diseaseAreaStrategies: IMetadataOption[]; // Unique DAS values from SPECTRA_DAS list
  diseaseAreaStrategyRelationships: IDiseaseAreaStrategyRelationship[]; // TA -> DAS rows from SPECTRA_DAS list
  therapeuticAreas: IMetadataOption[]; // From SPECTRA_TherapeuticAreas list
  subTherapeuticAreas: IMetadataOption[]; // From SPECTRA_SubTherapeuticAreas list
  indications: IMetadataOption[]; // From SPECTRA_Indications list
  lineOfTherapy: IMetadataOption[]; // From SPECTRA_LineOfTherapy list
  projectPaidRelationships: IProjectPaidRelationship[]; // From SPECTRA_ProjectPAID list
}

// ─────────────────────────────────────────────────────────────────
// Empty state — used as initial value before data loads
// ─────────────────────────────────────────────────────────────────
export const EMPTY_METADATA: IMetadataOptions = {
  assets: [],
  documentTypes: [],
  diseaseAreaStrategies: [],
  diseaseAreaStrategyRelationships: [],
  therapeuticAreas: [],
  subTherapeuticAreas: [],
  indications: [],
  lineOfTherapy: [],
  projectPaidRelationships: [],
};

// ─────────────────────────────────────────────────────────────────
// Helper — extract plain string values from IMetadataOption[]
// Used when mapping options to dropdown items:
//   toValueArray(options.assets).map(a => ({ key: a, text: a }))
// ─────────────────────────────────────────────────────────────────
export const toValueArray = (options: IMetadataOption[]): string[] =>
  Array.from(
    new Set(
      options
        .map((o) => (typeof o.value === "string" ? o.value.trim() : ""))
        .filter((value) => value.length > 0),
    ),
  );
