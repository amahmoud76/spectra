import { IMetadataOption } from "../interfaces/IMetadataOptions";

/**
 * Build search tokens for a document based on its selected
 * asset, therapeuticArea, indication, and documentType values.
 *
 * Looks up each selected value in the metadata options and
 * collects all associated searchTokens. These are stored on
 * the document's SPECTRASearchTokens field so that synonym
 * search works without needing the metadata options at query time.
 *
 * Called by UploadPanel and ReplacePanel before submission.
 */
export const buildDocumentSearchTokens = (
  selectedDocumentTypes: string[],
  selectedAssets: string[],
  selectedTAs: string[],
  selectedIndications: string[],
  documentTypeOptions: IMetadataOption[],
  assetOptions: IMetadataOption[],
  taOptions: IMetadataOption[],
  indicationOptions: IMetadataOption[],
): string[] => {
  const tokens = new Set<string>();

  // Collect tokens from selected document types
  selectedDocumentTypes.forEach((docTypeValue) => {
    const option = documentTypeOptions.find((o) => o.value === docTypeValue);
    if (option) {
      option.searchTokens.forEach((t) => tokens.add(t.toLowerCase()));
    }
  });

  // Collect tokens from selected assets
  selectedAssets.forEach((assetValue) => {
    const option = assetOptions.find((o) => o.value === assetValue);
    if (option) {
      option.searchTokens.forEach((t) => tokens.add(t.toLowerCase()));
    }
  });

  // Collect tokens from selected TAs
  selectedTAs.forEach((taValue) => {
    const option = taOptions.find((o) => o.value === taValue);
    if (option) {
      option.searchTokens.forEach((t) => tokens.add(t.toLowerCase()));
    }
  });

  // Collect tokens from selected indications
  selectedIndications.forEach((indValue) => {
    const option = indicationOptions.find((o) => o.value === indValue);
    if (option) {
      option.searchTokens.forEach((t) => tokens.add(t.toLowerCase()));
    }
  });

  return Array.from(tokens);
};
