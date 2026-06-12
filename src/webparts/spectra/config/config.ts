// ─────────────────────────────────────────────────────────────────
// SPECTRA Document Repository — Configuration
//
// This is the ONLY file you touch when switching between
// mock data and live SharePoint data.
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
// MOCK CONFIGURATION
// Set USE_MOCK = false when SharePoint lists and library are ready
// ─────────────────────────────────────────────────────────────────
export const USE_MOCK = false;

// Change to simulate different role behaviors during development:
//   'admin'       → sees all documents, all actions available
//   'contributor'  → sees current docs, upload + archive & replace
//   'viewer'       → sees current docs, read-only
export const MOCK_ROLE: "admin" | "contributor" | "viewer" = "admin";

// ─────────────────────────────────────────────────────────────────
// PAGINATION
// ─────────────────────────────────────────────────────────────────
export const PAGE_SIZE_DEFAULT = 25;
export const PAGE_SIZE_OPTIONS = [25, 50, 75];

// Inactivity lock timeout (minutes)
export const INACTIVITY_TIMEOUT_DEFAULT_MINUTES = 30;
export const INACTIVITY_TIMEOUT_OPTIONS_MINUTES = [1, 2, 10, 15, 30, 45, 60];
export const INACTIVITY_TEST_MODE_MAX_MINUTES = 2;

// ─────────────────────────────────────────────────────────────────
// CACHE
// MetadataService caches reference list data in localStorage
// ─────────────────────────────────────────────────────────────────
export const METADATA_CACHE_KEY = "spectra_metadata_v2";
export const METADATA_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

export const AUTH_CACHE_KEY_PREFIX = "spectra_auth_v2_";
export const AUTH_CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

// Auth decision tracing (console logs)
// Set to false to disable verbose auth role resolution logs.
export const AUTH_DEBUG_LOGS = true;
export const AUTH_LOG_PREFIX = "[SPECTRA AUTH]";

export const HEADER_CONFIG_CACHE_KEY = "spectra_header_config";
export const HEADER_CONFIG_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

// ─────────────────────────────────────────────────────────────────
// COMMENTS
// ─────────────────────────────────────────────────────────────────
export const COMMENTS_MAX_LENGTH = 100;

// ─────────────────────────────────────────────────────────────────
// NOTIFICATION
// Success banners auto-dismiss after this duration
// Error banners persist until manually dismissed
// ─────────────────────────────────────────────────────────────────
export const NOTIFICATION_AUTO_DISMISS_MS = 6000; // 6 seconds
export const GENERIC_ERROR_MESSAGE = "An error occurred and has been logged.";

// ─────────────────────────────────────────────────────────────────
// SEARCH
// Debounce delay for search input
// ─────────────────────────────────────────────────────────────────
export const SEARCH_DEBOUNCE_MS = 300;

// ─────────────────────────────────────────────────────────────────
// FILE UPLOAD
// ─────────────────────────────────────────────────────────────────
// SharePoint REST API Files/add endpoint supports up to ~250 MB.
// Files above this fail with HTTP 400 on upload.
export const MAX_FILE_SIZE_BYTES = 250 * 1024 * 1024; // 250 MB
export const MAX_FILE_SIZE_DISPLAY = "250 MB";

export const ACCEPTED_FILE_EXTENSIONS = [".pdf", ".docx", ".pptx", ".xlsx"];
export const ACCEPTED_FILE_TYPES = [
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
];
export const ACCEPTED_FILE_DISPLAY = "PDF, Word, XLSX, PPT";

// ─────────────────────────────────────────────────────────────────
// DOCUMENT TYPES (BRD Appendix A — Release 1)
// ─────────────────────────────────────────────────────────────────
export const DOCUMENT_TYPES = [
  "R&D One-Pager",
  "PDS",
  "DAS",
  "TPP",
  "TPC",
  "IEP",
  "EBP",
  "EIVP",
  "IAS",
] as const;

export type DocumentType = (typeof DOCUMENT_TYPES)[number];

// R&D One-Pager special case (BRD 3.3):
// When this document type is selected, all metadata fields
// except Document Type become hidden/optional
export const ONE_PAGER_DOC_TYPE = "R&D One-Pager";

// Document types visible to admins only (hidden from contributor/viewer dropdowns)
export const ADMIN_ONLY_DOC_TYPES: string[] = ["PDS", "R&D One-Pager"];

// ─────────────────────────────────────────────────────────────────
// THERAPEUTIC AREAS
// General Medicine is EXCLUDED per business decision
// ─────────────────────────────────────────────────────────────────
export const THERAPEUTIC_AREAS = [
  "Aesthetics",
  "Eye Care",
  "Immunology",
  "Neuroscience",
  "Oncology",
  "Specialty",
] as const;

export type TherapeuticArea = (typeof THERAPEUTIC_AREAS)[number];

// ─────────────────────────────────────────────────────────────────
// LINE OF THERAPY
// ─────────────────────────────────────────────────────────────────
export const LINE_OF_THERAPY_OPTIONS = [
  "1L",
  "2L",
  "2L+",
  "2L+ RR",
  "3L",
  "3L+",
] as const;

// ─────────────────────────────────────────────────────────────────
// DOCUMENT STATUS
// ─────────────────────────────────────────────────────────────────
export const DOCUMENT_STATUSES = ["Current", "Archive"] as const;
export type DocumentStatus = (typeof DOCUMENT_STATUSES)[number];

// ─────────────────────────────────────────────────────────────────
// FIREWALL ASSETS (BRD Appendix E)
// Uploads for these assets are BLOCKED in Release 1.
// Display error and prevent upload if user selects one.
// ─────────────────────────────────────────────────────────────────
export interface IFirewallAsset {
  therapeuticArea: string;
  assetName: string;
}

export const FIREWALL_ASSETS: IFirewallAsset[] = [
  { therapeuticArea: "Neuroscience", assetName: "Vraylar External" },
  { therapeuticArea: "Neuroscience", assetName: "Vraylar Internal" },
  { therapeuticArea: "Oncology", assetName: "Surzetoclax" },
  { therapeuticArea: "Oncology", assetName: "Venetoclax" },
  { therapeuticArea: "Oncology", assetName: "Imbruvica" },
  { therapeuticArea: "Oncology", assetName: "ABBV-101" },
];

// All Toxins assets are firewalled — check TA = "Toxins"
export const FIREWALL_TA_BLOCKED = ["Toxins"];

// Helper: check if an asset is firewalled
export const isFirewalledAsset = (
  assetName: string,
  therapeuticArea?: string,
): boolean => {
  // Check if the entire TA is blocked
  if (therapeuticArea && FIREWALL_TA_BLOCKED.includes(therapeuticArea)) {
    return true;
  }
  // Check individual firewalled assets (case-insensitive)
  return FIREWALL_ASSETS.some(
    (fa) => fa.assetName.toLowerCase() === assetName.toLowerCase(),
  );
};

// ─────────────────────────────────────────────────────────────────
// FILE NAMING CONVENTION (BRD Appendix H)
//
// Format:
//   Base order: DocType-TA-Asset-Indication[-SubTA][-LineOfTherapy].ext
//
// Rules:
//   - Multiple values concatenated with '+' in alphabetical order
//   - Asset is placed before Indication
//   - Sub-TA included for therapeutic areas listed in TA_INCLUDES_SUB_TA
//   - Line of Therapy included if it exists
//   - Spaces replaced with hyphens in each segment
//   - Effective Date excluded from filename (Release 1); stored as metadata only
//   - DAS: DocType-TA[-SubTA]-DiseaseArea only (Disease Area required, not optional)
// ─────────────────────────────────────────────────────────────────

// Abbreviation maps for file naming
// Document types use their abbreviation in the filename
export const DOC_TYPE_ABBREVIATIONS: Record<string, string> = {
  "R&D One-Pager": "OnePager",
  PDS: "PDS",
  DAS: "DAS",
  TPP: "TPP",
  TPC: "TPC",
  IEP: "IEP",
  EBP: "EBP",
  EIVP: "EIVP",
  IAS: "IAS",
};

// Therapeutic area abbreviations for file naming
export const TA_ABBREVIATIONS: Record<string, string> = {
  Aesthetics: "AES",
  "Eye Care": "EYE",
  Immunology: "IMM",
  Neuroscience: "NEU",
  Oncology: "ONC",
  Specialty: "SPEC",
};

// TAs that include Sub-TA in the filename
export const TA_INCLUDES_SUB_TA = ["Aesthetics", "Oncology"];

// TAs that commonly include Line of Therapy in the filename
// (Line of Therapy is included if it exists, regardless of TA,
//  but it is primarily used for these TAs)
export const TA_PRIMARY_LOT = ["Oncology"];

// ─────────────────────────────────────────────────────────────────
// SHAREPOINT LIST NAMES
// Used by MetadataService to fetch reference data
// ─────────────────────────────────────────────────────────────────
export const SP_LIST_NAMES = {
  assets: "SPECTRA_Assets",
  documentTypes: "SPECTRA_DocumentTypes",
  diseaseAreaStrategies: "SPECTRA_DAS",
  therapeuticAreas: "SPECTRA_TherapeuticAreas",
  subTherapeuticAreas: "SPECTRA_SubTherapeuticAreas",
  indications: "SPECTRA_Indications",
  lineOfTherapy: "SPECTRA_LineOfTherapy",
  projectPaid: "SPECTRA_ProjectPAID",
  headerSettings: "SPECTRA_HeaderSettings",
  headerLinks: "SPECTRA_HeaderLinks",
} as const;

// Document library name
export const SP_DOCUMENT_LIBRARY = "SPECTRA_Documents";

// ─────────────────────────────────────────────────────────────────
// SHAREPOINT ROLE GROUPS (dynamic AD-group source)
// AuthService queries these SharePoint groups at runtime and pulls
// the AD groups (PrincipalType = 4) assigned to each, so admins can
// add/remove AD groups in SharePoint without redeploying.
// ─────────────────────────────────────────────────────────────────
export const SP_ROLE_GROUPS = {
  admins: "Document Repository Owners",
  contributors: "Document Repository Members",
  viewers: "Document Repository Visitors",
} as const;

// ─────────────────────────────────────────────────────────────────
// ENTRA ID / AD GROUP NAMES (fallback only)
// Used by AuthService when the dynamic SharePoint role-group lookup
// fails or returns empty. The dynamic SP_ROLE_GROUPS query is the
// primary source of truth.
// ─────────────────────────────────────────────────────────────────
export const AD_GROUPS = {
  admins: ["APP-INFORM-PipelineAnalyticsOps", "APP-QlikSense-PPM-BTS"],
  contributors: ["APP-IlluminationHub-Contributor"],
  viewers: [
    "APP-IlluminationHub-AdHoc",
    "APP-QlikSense-PMO-Roster-PROD",
    "APP-IlluminationHub-LT",
  ],
} as const;

// When false, AuthService will not call Microsoft Graph for group checks.
// It will use MOCK_GROUP_MEMBERSHIP first, then fall back to MOCK_ROLE.
export const USE_AD_GROUPS = true;

// Mock group membership by email for non-AD environments.
// Keys must be lowercase email addresses.
export const MOCK_GROUP_MEMBERSHIP: Record<string, (keyof typeof AD_GROUPS)[]> =
  {
    "ahmedm@slalom.com": ["admins"],
  };

// ─────────────────────────────────────────────────────────────────
// MULTI-VALUE FIELD SEPARATOR
// Used when reading/writing multi-value fields to/from SharePoint
// SharePoint stores: "CLL; MCL; WM"
// Application uses: ["CLL", "MCL", "WM"]
// ─────────────────────────────────────────────────────────────────
export const MULTI_VALUE_SEPARATOR = "; ";

// Search tokens use comma separator (different from other multi-value fields)
export const SEARCH_TOKENS_SEPARATOR = ", ";

// ─────────────────────────────────────────────────────────────────
// HELP
// ─────────────────────────────────────────────────────────────────
export const HELP_EMAIL = "spectra-support@abbvie.com";
export const HELP_GUIDE_URL = ""; // To be configured with actual URL
