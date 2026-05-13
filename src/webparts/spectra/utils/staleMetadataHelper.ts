import {
  IMetadataOptions,
  toValueArray,
} from "../interfaces/IMetadataOptions";
import {
  getAllPaids,
  getCascadedOptionsMulti,
  getDiseaseAreaStrategiesForTherapeuticArea,
} from "./cascadingFilterHelper";

export interface IMetadataSelectionState {
  documentType: string;
  therapeuticArea: string[];
  subTherapeuticArea: string[];
  diseaseArea: string[];
  asset: string[];
  indication: string[];
  lineOfTherapy: string[];
  paid: string[];
}

export interface IFieldStaleValues {
  documentType: string[];
  therapeuticArea: string[];
  subTherapeuticArea: string[];
  diseaseArea: string[];
  asset: string[];
  indication: string[];
  lineOfTherapy: string[];
  paid: string[];
}

export interface ICascadeInvalidValues {
  subTherapeuticArea: string[];
  diseaseArea: string[];
  indication: string[];
  asset: string[];
  paid: string[];
}

export interface IStaleMetadataState {
  staleFromSource: IFieldStaleValues;
  cascadeInvalid: ICascadeInvalidValues;
  hasStaleValues: boolean;
  summaryMessage: string;
}

const EMPTY_FIELD_STALE_VALUES: IFieldStaleValues = {
  documentType: [],
  therapeuticArea: [],
  subTherapeuticArea: [],
  diseaseArea: [],
  asset: [],
  indication: [],
  lineOfTherapy: [],
  paid: [],
};

const EMPTY_CASCADE_INVALID_VALUES: ICascadeInvalidValues = {
  subTherapeuticArea: [],
  diseaseArea: [],
  indication: [],
  asset: [],
  paid: [],
};

const hasAnyValues = <T extends { [K in keyof T]: string[] }>(obj: T): boolean =>
  (Object.keys(obj) as Array<keyof T>).some((key) => obj[key].length > 0);

const unique = (values: string[]): string[] =>
  Array.from(new Set(values.filter(Boolean)));

const diff = (selected: string[], allowed: string[]): string[] => {
  const allowedSet = new Set(allowed);
  return unique(selected.filter((value) => !allowedSet.has(value)));
};

const buildSummaryMessage = (
  staleFromSource: IFieldStaleValues,
  cascadeInvalid: ICascadeInvalidValues,
): string => {
  const parts: string[] = [];

  const addPart = (label: string, values: string[]): void => {
    if (values.length === 0) return;
    parts.push(`${label}: ${values.join(", ")}`);
  };

  addPart("Document Type", staleFromSource.documentType);
  addPart("Therapeutic Area", staleFromSource.therapeuticArea);
  addPart("Sub-Therapeutic Area", staleFromSource.subTherapeuticArea);
  addPart("Disease Area", staleFromSource.diseaseArea);
  addPart("Asset", staleFromSource.asset);
  addPart("Indication", staleFromSource.indication);
  addPart("Line of Therapy", staleFromSource.lineOfTherapy);
  addPart("PAID", staleFromSource.paid);

  addPart("Cascade mismatch - Sub-Therapeutic Area", cascadeInvalid.subTherapeuticArea);
  addPart("Cascade mismatch - Disease Area", cascadeInvalid.diseaseArea);
  addPart("Cascade mismatch - Indication", cascadeInvalid.indication);
  addPart("Cascade mismatch - Asset", cascadeInvalid.asset);
  addPart("Cascade mismatch - PAID", cascadeInvalid.paid);

  if (parts.length === 0) return "";

  return `Some saved metadata values are no longer aligned with current reference metadata. ${parts.join(" | ")}`;
};

export const getCascadeInvalidValues = (
  selection: IMetadataSelectionState,
  options: IMetadataOptions,
): ICascadeInvalidValues => {
  const selectedTherapeuticArea = selection.therapeuticArea[0] || "";

  const subTherapeuticAreaOptions = selectedTherapeuticArea
    ? options.subTherapeuticAreas
        .filter((option) =>
          option.searchTokens.some(
            (token) =>
              token.trim().toLowerCase() ===
              selectedTherapeuticArea.trim().toLowerCase(),
          ),
        )
        .map((option) => option.value)
    : toValueArray(options.subTherapeuticAreas);

  const diseaseAreaOptions = getDiseaseAreaStrategiesForTherapeuticArea(
    options.diseaseAreaStrategyRelationships,
    selectedTherapeuticArea || undefined,
  );

  const cascadedIndications = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: selectedTherapeuticArea || undefined,
      selectedAssets: selection.asset[0] || undefined,
      selectedPaids: selection.paid.length > 0 ? selection.paid : undefined,
    }
  );

  const cascadedAssets = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: selectedTherapeuticArea || undefined,
      selectedIndications:
        selection.indication.length > 0 ? selection.indication : undefined,
      selectedPaids: selection.paid.length > 0 ? selection.paid : undefined,
    }
  );

  const cascadedPaids = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: selectedTherapeuticArea || undefined,
      selectedAssets: selection.asset[0] || undefined,
      selectedIndications:
        selection.indication.length > 0 ? selection.indication : undefined,
    }
  );

  return {
    subTherapeuticArea: diff(
      selection.subTherapeuticArea,
      subTherapeuticAreaOptions,
    ),
    diseaseArea: diff(selection.diseaseArea, diseaseAreaOptions),
    indication: diff(
      selection.indication,
      cascadedIndications.availableIndications,
    ),
    asset: diff(selection.asset, cascadedAssets.availableAssets),
    paid: diff(selection.paid, cascadedPaids.availablePaids),
  };
};

export const getStaleMetadataState = (
  selection: IMetadataSelectionState,
  options: IMetadataOptions,
): IStaleMetadataState => {
  const staleFromSource: IFieldStaleValues = {
    ...EMPTY_FIELD_STALE_VALUES,
    documentType:
      selection.documentType &&
      !toValueArray(options.documentTypes).includes(selection.documentType)
        ? [selection.documentType]
        : [],
    therapeuticArea: diff(
      selection.therapeuticArea,
      toValueArray(options.therapeuticAreas),
    ),
    subTherapeuticArea: diff(
      selection.subTherapeuticArea,
      toValueArray(options.subTherapeuticAreas),
    ),
    diseaseArea: diff(
      selection.diseaseArea,
      toValueArray(options.diseaseAreaStrategies),
    ),
    asset: diff(selection.asset, toValueArray(options.assets)),
    indication: diff(selection.indication, toValueArray(options.indications)),
    lineOfTherapy: diff(
      selection.lineOfTherapy,
      toValueArray(options.lineOfTherapy),
    ),
    paid: diff(selection.paid, getAllPaids(options.projectPaidRelationships)),
  };

  const cascadeInvalid = getCascadeInvalidValues(selection, options);
  const summaryMessage = buildSummaryMessage(staleFromSource, cascadeInvalid);

  const hasStaleValues =
    hasAnyValues(staleFromSource) || hasAnyValues(cascadeInvalid);

  return {
    staleFromSource,
    cascadeInvalid,
    hasStaleValues,
    summaryMessage,
  };
};

export const hasNewInvalidValues = (
  currentInvalidValues: string[],
  baselineInvalidValues: string[],
): string[] => {
  const baselineSet = new Set(baselineInvalidValues);
  return currentInvalidValues.filter((value) => !baselineSet.has(value));
};

export const emptyCascadeInvalidValues = (): ICascadeInvalidValues => ({
  ...EMPTY_CASCADE_INVALID_VALUES,
});
