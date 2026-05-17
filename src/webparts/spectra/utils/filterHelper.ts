import { IDocument } from "../interfaces/IDocument";
import { IFilterState } from "../interfaces/IFilterState";
import { parseISO, isValid, isWithinInterval } from "date-fns";

export type SearchMatchKind = "exact" | "close";

export interface IApplyFiltersResult {
  documents: IDocument[];
  /** Per-document match quality when search text is active. */
  matchKindByDocumentId: Map<string, SearchMatchKind>;
  exactMatchCount: number;
  closeMatchCount: number;
  /** True when any close (fuzzy) matches are included in results. */
  usedFuzzySearch: boolean;
  /** True when both exact and close matches are present. */
  hasMixedSearchResults: boolean;
}

const tokenizeSearchText = (rawText: string): string[] =>
  rawText
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean);

const toAlphanumeric = (value: string): string =>
  value.toLowerCase().replace(/[^a-z0-9]/g, "");

const tokenizeAlphanumeric = (value: string): string[] =>
  value
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 0);

const levenshteinDistance = (a: string, b: string): number => {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;

  const rows = a.length + 1;
  const cols = b.length + 1;
  const matrix: number[][] = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => 0),
  );

  for (let i = 0; i < rows; i++) matrix[i][0] = i;
  for (let j = 0; j < cols; j++) matrix[0][j] = j;

  for (let i = 1; i < rows; i++) {
    for (let j = 1; j < cols; j++) {
      const substitutionCost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + substitutionCost,
      );
    }
  }

  return matrix[a.length][b.length];
};

/** Short codes (TPC, TPP, AES) — token-only fuzzy, at most one edit, similar length. */
const SHORT_FUZZY_TERM_MAX_LENGTH = 4;

const maxAllowedEditDistance = (length: number): number => {
  if (length <= 4) return 1;
  if (length <= 8) return 2;
  if (length <= 14) return 3;
  return 4;
};

const termHasDigits = (term: string): boolean => /\d/.test(term);

const isWithinFuzzyEditDistance = (a: string, b: string): boolean => {
  if (!a || !b) return false;
  if (a === b) return true;

  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;

  if (shorter.length < 3) {
    return false;
  }

  // Short query tokens: one typo only, similar-length counterpart (e.g. TPP ↔ TPC, not IEP or ABT).
  if (shorter.length <= SHORT_FUZZY_TERM_MAX_LENGTH) {
    if (longer.length - shorter.length > 1) {
      return false;
    }
    return levenshteinDistance(shorter, longer) <= 1;
  }

  if (a.includes(b) || b.includes(a)) {
    return true;
  }

  if (longer.length - shorter.length > 2) {
    return false;
  }

  const distance = levenshteinDistance(shorter, longer);
  return distance <= maxAllowedEditDistance(shorter.length);
};

const matchesTokenFuzzy = (term: string, candidateToken: string): boolean => {
  const normalizedTerm = toAlphanumeric(term);
  const normalizedCandidate = toAlphanumeric(candidateToken);

  if (!normalizedTerm || !normalizedCandidate) return false;
  if (normalizedCandidate.includes(normalizedTerm)) return true;

  const digitsFromTerm = normalizedTerm.replace(/\D/g, "");
  const digitsFromCandidate = normalizedCandidate.replace(/\D/g, "");

  if (termHasDigits(normalizedTerm)) {
    if (!digitsFromCandidate) return false;

    if (
      digitsFromTerm &&
      (digitsFromCandidate.includes(digitsFromTerm) ||
        digitsFromTerm.includes(digitsFromCandidate))
    ) {
      return true;
    }

    if (
      digitsFromTerm.length >= 2 &&
      digitsFromCandidate.length >= 2 &&
      levenshteinDistance(digitsFromTerm, digitsFromCandidate) <= 1
    ) {
      return true;
    }
  }

  return isWithinFuzzyEditDistance(normalizedTerm, normalizedCandidate);
};

/**
 * Fuzzy-match a query term against one corpus field.
 * Short codes match delimited tokens only (no sliding windows inside long filenames).
 * Longer terms may also match the full normalized field value (e.g. disease names).
 */
const matchesTermFuzzyInValue = (term: string, rawValue: string): boolean => {
  if (!term || !rawValue) return false;

  const value = rawValue.toLowerCase();

  if (term.length < 2) {
    return value.includes(term);
  }

  const tokens = tokenizeAlphanumeric(value);
  if (tokens.some((token) => matchesTokenFuzzy(term, token))) {
    return true;
  }

  const normalizedTerm = toAlphanumeric(term);

  // Whole-field fuzzy only for longer words (typos in titles/descriptions), not 3-letter codes.
  if (normalizedTerm.length > SHORT_FUZZY_TERM_MAX_LENGTH) {
    const normalizedValue = toAlphanumeric(value);
    if (
      normalizedValue.length >= normalizedTerm.length &&
      isWithinFuzzyEditDistance(normalizedTerm, normalizedValue)
    ) {
      return true;
    }
  }

  return false;
};

const matchesSearchTermsStrict = (
  searchCorpus: string[],
  queryTerms: string[],
): boolean =>
  queryTerms.length === 0 ||
  queryTerms.every((term) =>
    searchCorpus.some((value) => value.includes(term)),
  );

const matchesSearchTermsFuzzy = (
  searchCorpus: string[],
  queryTerms: string[],
): boolean =>
  queryTerms.length === 0 ||
  queryTerms.every((term) =>
    searchCorpus.some((value) => matchesTermFuzzyInValue(term, value)),
  );

const buildDocumentSearchCorpus = (doc: IDocument): string[] => {
  const scalarValues: string[] = [
    doc.fileName,
    doc.fileExtension,
    doc.fileUrl,
    doc.createdBy,
    doc.createdByEmail,
    doc.modifiedBy,
    doc.modifiedByEmail,
    doc.uploadDate,
    doc.title,
    doc.documentType,
    doc.effectiveDate,
    doc.status,
    doc.description,
    doc.comments,
    doc.spectra511,
    doc.immutableFileName,
  ];

  const arrayValues: string[] = [
    ...doc.asset,
    ...doc.therapeuticArea,
    ...doc.subTherapeuticArea,
    ...doc.indication,
    ...doc.lineOfTherapy,
    ...doc.paid,
    ...doc.diseaseArea,
    ...(doc.searchTokens || []),
  ];

  return Array.from(
    new Set(
      [...scalarValues, ...arrayValues]
        .map((value) => value?.toString().trim().toLowerCase())
        .filter(Boolean),
    ),
  );
};

/**
 * Apply all active filters + search text to a document array.
 * Returns a new filtered array — does not mutate the input.
 * Filters combine with AND logic.
 *
 * Uses date-fns for timezone-safe date parsing and range checking.
 */
export const applyFilters = (
  documents: IDocument[],
  filters: IFilterState,
): IDocument[] => {
  return applyFiltersWithMeta(documents, filters).documents;
};

export const applyFiltersWithMeta = (
  documents: IDocument[],
  filters: IFilterState,
): IApplyFiltersResult => {
  const matchesNonSearchFilters = (doc: IDocument): boolean => {
    // ── Document Type filter ──────────────────────────────
    if (filters.documentType.length > 0) {
      if (!filters.documentType.includes(doc.documentType)) return false;
    }

    // ── Therapeutic Area filter ────────────────────────────
    if (filters.therapeuticArea.length > 0) {
      if (
        !doc.therapeuticArea.some((ta) => filters.therapeuticArea.includes(ta))
      )
        return false;
    }

    // ── Sub-Therapeutic Area filter ────────────────────────
    if (filters.subTherapeuticArea.length > 0) {
      if (
        !doc.subTherapeuticArea.some((sta) =>
          filters.subTherapeuticArea.includes(sta),
        )
      )
        return false;
    }

    // ── Asset filter ──────────────────────────────────────
    if (filters.asset.length > 0) {
      if (!doc.asset.some((a) => filters.asset.includes(a))) return false;
    }

    // ── Indication filter ─────────────────────────────────
    if (filters.indication.length > 0) {
      if (!doc.indication.some((ind) => filters.indication.includes(ind)))
        return false;
    }

    // ── Line of Therapy filter ────────────────────────────
    if (filters.lineOfTherapy.length > 0) {
      if (!doc.lineOfTherapy.some((lot) => filters.lineOfTherapy.includes(lot)))
        return false;
    }

    // ── PAID filter ───────────────────────────────────────
    if (filters.paid.length > 0) {
      if (!doc.paid.some((p) => filters.paid.includes(p))) return false;
    }

    // ── Disease Area filter ───────────────────────────────
    if (filters.diseaseArea.length > 0) {
      if (!doc.diseaseArea.some((da) => filters.diseaseArea.includes(da)))
        return false;
    }

    // ── Effective Date range (date-fns — timezone safe) ───
    if (filters.effectiveDateFrom || filters.effectiveDateTo) {
      const docDate = parseISO(doc.effectiveDate);
      if (!isValid(docDate)) return false;

      if (filters.effectiveDateFrom && filters.effectiveDateTo) {
        if (
          !isWithinInterval(docDate, {
            start: filters.effectiveDateFrom,
            end: filters.effectiveDateTo,
          })
        )
          return false;
      } else if (
        filters.effectiveDateFrom &&
        docDate < filters.effectiveDateFrom
      ) {
        return false;
      } else if (filters.effectiveDateTo && docDate > filters.effectiveDateTo) {
        return false;
      }
    }

    // ── Upload Date range (date-fns — timezone safe) ──────
    if (filters.uploadDateFrom || filters.uploadDateTo) {
      const uploadDate = parseISO(doc.uploadDate);
      if (!isValid(uploadDate)) return false;
      if (filters.uploadDateFrom && uploadDate < filters.uploadDateFrom)
        return false;
      if (filters.uploadDateTo && uploadDate > filters.uploadDateTo)
        return false;
    }

    return true;
  };

  const baseFiltered = documents.filter(matchesNonSearchFilters);

  if (!filters.searchText.trim()) {
    return {
      documents: baseFiltered,
      matchKindByDocumentId: new Map(),
      exactMatchCount: 0,
      closeMatchCount: 0,
      usedFuzzySearch: false,
      hasMixedSearchResults: false,
    };
  }

  const queryTerms = tokenizeSearchText(filters.searchText);
  const strictMatches = baseFiltered.filter((doc) =>
    matchesSearchTermsStrict(buildDocumentSearchCorpus(doc), queryTerms),
  );
  const strictIds = new Set(strictMatches.map((doc) => doc.id));

  const closeOnlyMatches = baseFiltered.filter((doc) => {
    if (strictIds.has(doc.id)) return false;
    return matchesSearchTermsFuzzy(
      buildDocumentSearchCorpus(doc),
      queryTerms,
    );
  });

  const rankedDocuments = [...strictMatches, ...closeOnlyMatches];
  const matchKindByDocumentId = new Map<string, SearchMatchKind>();
  strictMatches.forEach((doc) =>
    matchKindByDocumentId.set(doc.id, "exact"),
  );
  closeOnlyMatches.forEach((doc) =>
    matchKindByDocumentId.set(doc.id, "close"),
  );

  return {
    documents: rankedDocuments,
    matchKindByDocumentId,
    exactMatchCount: strictMatches.length,
    closeMatchCount: closeOnlyMatches.length,
    usedFuzzySearch: closeOnlyMatches.length > 0,
    hasMixedSearchResults:
      strictMatches.length > 0 && closeOnlyMatches.length > 0,
  };
};

/**
 * Get unique values from a single-value field across all documents.
 * Used to build filter dropdown options from the current document set.
 */
export const getUniqueValues = (
  documents: IDocument[],
  field: keyof Pick<
    IDocument,
    "documentType" | "lineOfTherapy" | "status"
  >,
): string[] =>
  Array.from(
    new Set(documents.map((doc) => doc[field] as string).filter(Boolean)),
  ).sort();

/**
 * Get unique values from a multi-value array field across all documents.
 */
export const getUniqueArrayValues = (
  documents: IDocument[],
  field: keyof Pick<
    IDocument,
    "asset" | "therapeuticArea" | "subTherapeuticArea" | "indication" | "paid"
  >,
): string[] =>
  Array.from(
    new Set(documents.flatMap((doc) => doc[field] as string[])),
  ).sort();

/**
 * Count the number of active filters (non-empty fields).
 */
export const countActiveFilters = (filters: IFilterState): number => {
  let count = 0;
  if (filters.documentType.length > 0) count++;
  if (filters.therapeuticArea.length > 0) count++;
  if (filters.subTherapeuticArea.length > 0) count++;
  if (filters.asset.length > 0) count++;
  if (filters.indication.length > 0) count++;
  if (filters.lineOfTherapy.length > 0) count++;
  if (filters.paid.length > 0) count++;
  if (filters.diseaseArea.length > 0) count++;
  if (filters.effectiveDateFrom || filters.effectiveDateTo) count++;
  if (filters.uploadDateFrom || filters.uploadDateTo) count++;
  return count;
};

export const hasActiveSearchText = (filters: IFilterState): boolean =>
  filters.searchText.trim().length > 0;
