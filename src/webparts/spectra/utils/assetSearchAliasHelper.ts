import { ISearchTokenRow } from "../interfaces/IMetadataOptions";

/**
 * Build an asset -> synonym map from SPECTRA_SearchTokens rows, keyed by
 * ASSET_NUMBER (the same string used as the Asset dropdown's option value).
 * Rows sharing an asset number have their SEARCH_TOKEN keywords merged, so
 * typing any synonym (e.g. "ABBV-132" or "emraclidine") matches that asset.
 */
export const buildAssetSearchAliases = (
  rows: ISearchTokenRow[],
): Record<string, string[]> => {
  const map: Record<string, string[]> = {};
  rows.forEach((row) => {
    const asset = row.assetNumber.trim();
    if (!asset) return;
    const existing = map[asset] ?? [];
    map[asset] = Array.from(new Set([...existing, ...row.searchTokens]));
  });
  return map;
};

/**
 * Display-only label for a document's Asset value: the asset ID(s) plus
 * any SPECTRA_SearchTokens synonyms, e.g. "ABBV-132 - emraclidine". Falls
 * back to just the asset ID(s) when there are no synonyms. Used by
 * DataTable and TilesView — never the underlying saved/cascaded value.
 */
export const formatAssetWithSynonyms = (
  assetValues: string[],
  assetSearchTokens?: string[],
): string => {
  const assetLabel = assetValues.join("; ");
  const tokens = (assetSearchTokens || []).join(" - ");
  return tokens ? `${assetLabel} - ${tokens}` : assetLabel;
};
