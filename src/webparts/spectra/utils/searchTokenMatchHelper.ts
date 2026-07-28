import { IDocument } from "../interfaces/IDocument";
import { ISearchTokenRow } from "../interfaces/IMetadataOptions";

// ─────────────────────────────────────────────────────────────────
// Search-token matching (SPECTRA_SearchTokens list)
//
// When the user's search text matches a SEARCH_TOKEN keyword, the
// search returns documents whose metadata satisfies the row's
// combination. Semantics (confirmed with product):
//   • AND within a row  — a document must match ALL non-empty fields
//     of a row.
//   • OR across rows    — the same keyword may appear in multiple
//     rows; a document matches if it satisfies ANY of those rows.
// Blank fields in a row are ignored.
// ─────────────────────────────────────────────────────────────────

/** Lowercase, trim, and collapse internal whitespace for comparison. */
const normalize = (value: string): string =>
  value.toLowerCase().trim().replace(/\s+/g, " ");

/** True when `value` (non-empty) is present in `arr` (case-insensitive). */
const arrayIncludes = (arr: string[], value: string): boolean => {
  const target = normalize(value);
  if (!target) return true; // blank row field → ignored (does not constrain)
  return arr.some((item) => normalize(item) === target);
};

/** Minimum query length before a partial (prefix) token match is allowed. */
const MIN_TOKEN_PREFIX_LENGTH = 3;

/**
 * Return all rows whose SEARCH_TOKEN keywords match the given search text.
 * The full (trimmed, normalized) search string is compared against each
 * keyword — multi-word tokens like "ABBV-706 + ZG006 (Zelgen)" are matched
 * as a whole, not tokenized on whitespace. A query of at least
 * MIN_TOKEN_PREFIX_LENGTH characters also matches keywords it is a prefix
 * of (e.g. "emr" matches "emraclidine"), mirroring the substring matching
 * already used by the asset filter panel.
 */
export const findMatchedTokenRows = (
  searchText: string,
  rows: ISearchTokenRow[],
): ISearchTokenRow[] => {
  const normalizedQuery = normalize(searchText);
  if (!normalizedQuery || !rows || rows.length === 0) return [];

  return rows.filter((row) =>
    row.searchTokens.some((token) => {
      const normalizedToken = normalize(token);
      if (normalizedToken === normalizedQuery) return true;
      return (
        normalizedQuery.length >= MIN_TOKEN_PREFIX_LENGTH &&
        normalizedToken.startsWith(normalizedQuery)
      );
    }),
  );
};

/**
 * True when the document corresponds to ANY of the supplied token rows.
 *
 * A SEARCH_TOKEN keyword is treated as an alias for the row's IDENTIFYING
 * columns — PROJECT_PAID and ASSET_NUMBER. A document matches when it
 * satisfies the row's non-empty identifying columns; the descriptive
 * columns (TA / Sub-TA / Indication / Disease Area / Line of Therapy) are
 * NOT required, so a document doesn't need to carry every column to be
 * returned by its token. If a row has neither identifying column populated,
 * it cannot match any document.
 */
export const docMatchesTokenRows = (
  doc: IDocument,
  rows: ISearchTokenRow[],
): boolean =>
  rows.some((row) => {
    const hasPaid = row.projectPaid.trim().length > 0;
    const hasAsset = row.assetNumber.trim().length > 0;
    if (!hasPaid && !hasAsset) return false;
    return (
      arrayIncludes(doc.paid, row.projectPaid) &&
      arrayIncludes(doc.asset, row.assetNumber)
    );
  });

/** True when the search text corresponds to at least one known token row. */
export const hasTokenMatch = (
  searchText: string,
  rows: ISearchTokenRow[],
): boolean => findMatchedTokenRows(searchText, rows).length > 0;
