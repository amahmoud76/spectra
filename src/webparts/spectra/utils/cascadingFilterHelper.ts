import {
  IDiseaseAreaStrategyRelationship,
  IProjectPaidRelationship,
} from "../interfaces/IMetadataOptions";

/**
 * Cascading metadata filter logic (BRD 3.4).
 *
 * Given the user's current selection and the full relationship data
 * from SPECTRA_ProjectPAID, returns the filtered options for each
 * dropdown field.
 *
 * Rules:
 *   - Select PAID → auto-populate TA, Asset, Indication
 *   - Select TA → filter to relevant PAIDs, Assets, Indications
 *   - Select Asset → filter to relevant TAs, PAIDs, Indications
 *   - Select Indication → filter to relevant TAs, Assets, PAIDs
 */

export interface ICascadeSelection {
  selectedPaid?: string;
  selectedTA?: string;
  selectedAsset?: string;
  selectedIndication?: string;
}

/**
 * Extended cascade selection supporting both single values and arrays.
 * When arrays are provided, uses Union (OR) logic for matching.
 */
export interface ICascadeSelectionMulti {
  selectedPaids?: string | string[];
  selectedTAs?: string | string[];
  selectedAssets?: string | string[];
  selectedIndications?: string | string[];
}

export interface ICascadeResult {
  availablePaids: string[];
  availableTAs: string[];
  availableAssets: string[];
  availableIndications: string[];
  availableSubTherapeuticAreas: string[];
  availableLineOfTherapies: string[];
}

/**
 * Deduplicate and sort a string array.
 */
function unique(values: string[]): string[] {
  return Array.from(new Set(values.filter(Boolean))).sort();
}

export const getCascadedOptions = (
  relationships: IProjectPaidRelationship[],
  selection: ICascadeSelection,
): ICascadeResult => {
  let filtered = [...relationships];

  // Apply each selection as a filter
  if (selection.selectedPaid) {
    filtered = filtered.filter((r) => r.projectPaid === selection.selectedPaid);
  }
  if (selection.selectedTA) {
    filtered = filtered.filter(
      (r) => r.therapeuticArea === selection.selectedTA,
    );
  }
  if (selection.selectedAsset) {
    filtered = filtered.filter(
      (r) => r.assetNumber === selection.selectedAsset,
    );
  }
  if (selection.selectedIndication) {
    filtered = filtered.filter(
      (r) => r.indication === selection.selectedIndication,
    );
  }

  // Extract unique values from the filtered set
  return {
    availablePaids: unique(filtered.map((r) => r.projectPaid)),
    availableTAs: unique(filtered.map((r) => r.therapeuticArea)),
    availableAssets: unique(filtered.map((r) => r.assetNumber)),
    availableIndications: unique(filtered.map((r) => r.indication)),
    availableSubTherapeuticAreas: unique(filtered.map((r) => r.subTherapeuticArea).filter((v): v is string => Boolean(v))),
    availableLineOfTherapies: unique(filtered.map((r) => r.lineOfTherapy).filter((v): v is string => Boolean(v))),
  };
};

/**
 * Get all unique PAID values from the relationship data.
 * Used to populate the PAID dropdown (no separate SPECTRA_PAIDs list needed).
 */
export const getAllPaids = (
  relationships: IProjectPaidRelationship[],
): string[] => unique(relationships.map((r) => r.projectPaid));

/**
 * Get unique DAS values, optionally filtered by Therapeutic Area.
 * This is intentionally separate from the ProjectPAID cascade.
 */
export const getDiseaseAreaStrategiesForTherapeuticArea = (
  relationships: IDiseaseAreaStrategyRelationship[],
  selectedTA?: string,
): string[] => {
  const normalizedTa = selectedTA?.trim().toLowerCase();

  if (!normalizedTa) {
    return unique(relationships.map((relationship) => relationship.value));
  }

  return unique(
    relationships
      .filter(
        (relationship) =>
          relationship.therapeuticArea.trim().toLowerCase() === normalizedTa,
      )
      .map((relationship) => relationship.value),
  );
};

  /**
   * Cascading options supporting Union/Permissive filtering.
   * When multiple values are selected in a dimension, returns options
   * that match ANY combination of those selections.
   *
   * Example: If user selects Indications [A, B] and PAIDs [1, 2],
   * will return Assets valid with ANY of: A×1, A×2, B×1, B×2
   */
  export const getCascadedOptionsMulti = (
    relationships: IProjectPaidRelationship[],
    selection: ICascadeSelectionMulti,
  ): ICascadeResult => {
    const toArray = (value: string | string[] | undefined): string[] => {
      if (!value) return [];
      return Array.isArray(value) ? value : [value];
    };

    const selectedPaidsArr = toArray(selection.selectedPaids);
    const selectedTAsArr = toArray(selection.selectedTAs);
    const selectedAssetsArr = toArray(selection.selectedAssets);
    const selectedIndicationsArr = toArray(selection.selectedIndications);

    let filtered = [...relationships];

    // Apply filters using OR logic within each dimension
    if (selectedPaidsArr.length > 0) {
      filtered = filtered.filter((r) => selectedPaidsArr.includes(r.projectPaid));
    }
    if (selectedTAsArr.length > 0) {
      filtered = filtered.filter((r) => selectedTAsArr.includes(r.therapeuticArea));
    }
    if (selectedAssetsArr.length > 0) {
      filtered = filtered.filter((r) => selectedAssetsArr.includes(r.assetNumber));
    }
    if (selectedIndicationsArr.length > 0) {
      filtered = filtered.filter((r) => selectedIndicationsArr.includes(r.indication));
    }

    return {
      availablePaids: unique(filtered.map((r) => r.projectPaid)),
      availableTAs: unique(filtered.map((r) => r.therapeuticArea)),
      availableAssets: unique(filtered.map((r) => r.assetNumber)),
      availableIndications: unique(filtered.map((r) => r.indication)),
      availableSubTherapeuticAreas: unique(filtered.map((r) => r.subTherapeuticArea).filter((v): v is string => Boolean(v))),
      availableLineOfTherapies: unique(filtered.map((r) => r.lineOfTherapy).filter((v): v is string => Boolean(v))),
    };
  };
