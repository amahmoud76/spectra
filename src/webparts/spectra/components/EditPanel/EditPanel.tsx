import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import {
  IMetadataOptions,
  toValueArray,
} from "../../interfaces/IMetadataOptions";
import { IUploadPayload } from "../../interfaces/IUploadPayload";
import {
  getDiseaseAreaStrategiesForTherapeuticArea,
  getCascadedOptionsMulti,
} from "../../utils/cascadingFilterHelper";
import { buildDocumentSearchTokens } from "../../utils/searchTokenHelper";
import { COMMENTS_MAX_LENGTH } from "../../config/config";
import {
  FIELD_NAMES,
  isFieldVisible,
  isFieldRequired,
  isFieldMultiSelect,
} from "../../config/fieldConfig";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { DatePicker } from "@fluentui/react/lib/DatePicker";
import { TextField } from "@fluentui/react/lib/TextField";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { ErrorBanner } from "../ErrorBanner/ErrorBanner";
import { SearchableDropdown } from "../SearchableDropdown/SearchableDropdown";
import { format, isValid, parse as parseDateFns, parseISO } from "date-fns";
import styles from "../SPECTRA.module.scss";

const FUTURE_EFFECTIVE_DATE_ERROR = "Effective Date cannot be in the future.";
const INVALID_CASCADE_SELECTION_ERROR =
  "One or more metadata selections are no longer valid for the current Therapeutic Area, Asset, Indication, or PAID values.";

interface ICascadeFieldErrors {
  subTherapeuticArea?: string;
  diseaseArea?: string;
  indication?: string;
  asset?: string;
  paid?: string;
}

function formatFieldLabels(labels: string[]): string {
  if (labels.length <= 1) return labels[0] || "";
  if (labels.length === 2) return `${labels[0]} and ${labels[1]}`;
  return `${labels.slice(0, -1).join(", ")}, and ${labels[labels.length - 1]}`;
}

function toPickerDate(iso: string): Date | undefined {
  if (!iso) return undefined;
  const parsed = parseISO(iso);
  return isValid(parsed) ? parsed : undefined;
}

function isFutureDate(date: Date): boolean {
  const candidate = new Date(date);
  candidate.setHours(0, 0, 0, 0);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return candidate.getTime() > today.getTime();
}

function normalizeCascadeValue(value: string): string {
  return value.trim().toLowerCase();
}

function keepCascadeOverlap(selected: string[], available: string[]): string[] {
  const availableByNormalized = new Map<string, string>();
  available.forEach((value) => {
    const trimmed = value.trim();
    const normalized = normalizeCascadeValue(trimmed);
    if (normalized && !availableByNormalized.has(normalized)) {
      availableByNormalized.set(normalized, trimmed);
    }
  });

  const dedupedSelected = Array.from(
    new Set(
      selected
        .map((value) => value.trim())
        .filter((value) => value.length > 0)
        .map((value) => normalizeCascadeValue(value)),
    ),
  );

  return dedupedSelected
    .filter((normalized) => availableByNormalized.has(normalized))
    .map((normalized) => availableByNormalized.get(normalized) || "")
    .filter((value) => value.length > 0);
}

function containsCascadeValue(values: string[], candidate: string): boolean {
  const normalizedCandidate = normalizeCascadeValue(candidate);
  return values.some(
    (value) => normalizeCascadeValue(value) === normalizedCandidate,
  );
}

export interface IEditPanelProps {
  isOpen: boolean;
  isAdmin: boolean;
  document: IDocument | null;
  options: IMetadataOptions;
  isSaving: boolean;
  onSave: (documentId: string, updates: Partial<IUploadPayload>) => void;
  onCancel: () => void;
  onArchiveClick: () => void;
  onDeleteClick: () => void;
  onReplaceFileClick: () => void;
  onReActivateClick: () => void;
}

export const EditPanel: React.FC<IEditPanelProps> = ({
  isOpen,
  isAdmin,
  document: doc,
  options,
  isSaving,
  onSave,
  onCancel,
  onArchiveClick,
  onDeleteClick,
  onReplaceFileClick,
  onReActivateClick,
}) => {
  // ── Form state ────────────────────────────────────────────
  const [documentType, setDocumentType] = React.useState("");
  const [therapeuticArea, setTherapeuticArea] = React.useState([] as string[]);
  const [subTherapeuticArea, setSubTherapeuticArea] = React.useState(
    [] as string[],
  );
  const [asset, setAsset] = React.useState([] as string[]);
  const [indication, setIndication] = React.useState([] as string[]);
  const [lineOfTherapy, setLineOfTherapy] = React.useState([] as string[]);
  const [paid, setPaid] = React.useState([] as string[]);
  const [diseaseArea, setDiseaseArea] = React.useState([] as string[]);
  const [effectiveDate, setEffectiveDate] = React.useState("");
  const [effectiveDateError, setEffectiveDateError] = React.useState("");
  const [comments, setComments] = React.useState("");
  const [cascadeErrors, setCascadeErrors] = React.useState<ICascadeFieldErrors>(
    {},
  );
  const [saveError, setSaveError] = React.useState("");
  const [constraintError, setConstraintError] = React.useState("");
  const topErrorRef = React.useRef<HTMLDivElement>(null);
  const panelBodyRef = React.useRef<HTMLDivElement>(null);

  const focusTopError = React.useCallback(() => {
    if (!topErrorRef.current) return;
    topErrorRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
    topErrorRef.current.focus();
  }, []);

  React.useEffect(() => {
    if (!saveError && !constraintError) return;
    // Wait for the banner to mount before scrolling/focusing.
    requestAnimationFrame(() => {
      focusTopError();
    });
  }, [saveError, constraintError, focusTopError]);

  // ── Pre-populate from document ────────────────────────────
  React.useEffect(() => {
    if (doc) {
      setDocumentType(doc.documentType);
      setTherapeuticArea(doc.therapeuticArea);
      setSubTherapeuticArea(doc.subTherapeuticArea);
      setAsset(doc.asset);
      setIndication(doc.indication);
      setLineOfTherapy(doc.lineOfTherapy);
      setPaid(doc.paid);
      setDiseaseArea(doc.diseaseArea ? [...doc.diseaseArea] : []);
      setEffectiveDate(doc.effectiveDate);
      setEffectiveDateError("");
      setCascadeErrors({});
      setComments(doc.comments);
      setSaveError("");
    }
  }, [doc]);

  const selectedTherapeuticArea =
    therapeuticArea.length > 0 ? therapeuticArea[0] : "";

  const fieldVisibility = React.useMemo(
    () => ({
      therapeuticArea: isFieldVisible(
        FIELD_NAMES.THERAPEUTIC_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
      subTherapeuticArea: isFieldVisible(
        FIELD_NAMES.SUB_THERAPEUTIC_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
      diseaseArea: isFieldVisible(
        FIELD_NAMES.DISEASE_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
      indication: isFieldVisible(
        FIELD_NAMES.INDICATION,
        documentType,
        selectedTherapeuticArea,
      ),
      lineOfTherapy: isFieldVisible(
        FIELD_NAMES.LINE_OF_THERAPY,
        documentType,
        selectedTherapeuticArea,
      ),
      asset: isFieldVisible(
        FIELD_NAMES.ASSET,
        documentType,
        selectedTherapeuticArea,
      ),
      paid: isFieldVisible(
        FIELD_NAMES.PAID,
        documentType,
        selectedTherapeuticArea,
      ),
      effectiveDate: isFieldVisible(
        FIELD_NAMES.EFFECTIVE_DATE,
        documentType,
        selectedTherapeuticArea,
      ),
      comments: isFieldVisible(
        FIELD_NAMES.COMMENTS,
        documentType,
        selectedTherapeuticArea,
      ),
    }),
    [documentType, selectedTherapeuticArea],
  );

  const fieldRequirements = React.useMemo(
    () => ({
      documentType: isFieldRequired(
        FIELD_NAMES.DOCUMENT_TYPE,
        documentType,
        selectedTherapeuticArea,
      ),
      therapeuticArea: isFieldRequired(
        FIELD_NAMES.THERAPEUTIC_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
      subTherapeuticArea: isFieldRequired(
        FIELD_NAMES.SUB_THERAPEUTIC_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
      diseaseArea: isFieldRequired(
        FIELD_NAMES.DISEASE_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
      indication: isFieldRequired(
        FIELD_NAMES.INDICATION,
        documentType,
        selectedTherapeuticArea,
      ),
      lineOfTherapy: isFieldRequired(
        FIELD_NAMES.LINE_OF_THERAPY,
        documentType,
        selectedTherapeuticArea,
      ),
      asset: isFieldRequired(
        FIELD_NAMES.ASSET,
        documentType,
        selectedTherapeuticArea,
      ),
      paid: isFieldRequired(
        FIELD_NAMES.PAID,
        documentType,
        selectedTherapeuticArea,
      ),
      effectiveDate: isFieldRequired(
        FIELD_NAMES.EFFECTIVE_DATE,
        documentType,
        selectedTherapeuticArea,
      ),
      comments: isFieldRequired(
        FIELD_NAMES.COMMENTS,
        documentType,
        selectedTherapeuticArea,
      ),
    }),
    [documentType, selectedTherapeuticArea],
  );

  const fieldMultiSelect = React.useMemo(
    () => ({
      indication: isFieldMultiSelect(
        FIELD_NAMES.INDICATION,
        documentType,
        selectedTherapeuticArea,
      ),
      paid: isFieldMultiSelect(
        FIELD_NAMES.PAID,
        documentType,
        selectedTherapeuticArea,
      ),
      diseaseArea: isFieldMultiSelect(
        FIELD_NAMES.DISEASE_AREA,
        documentType,
        selectedTherapeuticArea,
      ),
    }),
    [documentType, selectedTherapeuticArea],
  );

  // ── Clear fields when they become hidden ────────────────────
  React.useEffect(() => {
    if (!fieldVisibility.indication && indication.length > 0) {
      setIndication([]);
    }
    if (!fieldVisibility.asset && asset.length > 0) {
      setAsset([]);
    }
    if (!fieldVisibility.paid && paid.length > 0) {
      setPaid([]);
    }
    if (!fieldVisibility.diseaseArea && diseaseArea.length > 0) {
      setDiseaseArea([]);
    }
    if (!fieldVisibility.subTherapeuticArea && subTherapeuticArea.length > 0) {
      setSubTherapeuticArea([]);
    }
    if (!fieldVisibility.lineOfTherapy && lineOfTherapy.length > 0) {
      setLineOfTherapy([]);
    }
  }, [
    fieldVisibility,
    indication,
    asset,
    paid,
    diseaseArea,
    subTherapeuticArea,
    lineOfTherapy,
  ]);

  // ── Cascade option computations (for pruning use) ──────────

  const filteredDiseaseAreaStrategies = React.useMemo(
    () =>
      getDiseaseAreaStrategiesForTherapeuticArea(
        options.diseaseAreaStrategyRelationships,
        selectedTherapeuticArea || undefined,
      ),
    [options.diseaseAreaStrategyRelationships, selectedTherapeuticArea],
  );

  const filteredSubTherapeuticAreas = React.useMemo(() => {
    if (!selectedTherapeuticArea) {
      return toValueArray(options.subTherapeuticAreas);
    }

    const selectedTaNormalized = selectedTherapeuticArea.trim().toLowerCase();

    return options.subTherapeuticAreas
      .filter((option) =>
        option.searchTokens.some(
          (token) =>
            typeof token === "string" &&
            token.trim().toLowerCase() === selectedTaNormalized,
        ),
      )
      .map((option) =>
        typeof option.value === "string" ? option.value.trim() : "",
      )
      .filter((value) => value.length > 0)
      .filter((value, index, values) => values.indexOf(value) === index);
  }, [options.subTherapeuticAreas, selectedTherapeuticArea]);

  // ── Cascading — separate cascade per field ────────────────
  // Each field's cascade EXCLUDES its own selection to prevent
  // self-filtering (which would collapse the options list).

  // Cascade for Asset options — filter by TA, Indication, PAID (NOT Asset)
  const cascadeForAssets = React.useMemo(
    () =>
      getCascadedOptionsMulti(options.projectPaidRelationships, {
        selectedTAs: therapeuticArea,
        selectedIndications: indication,
        selectedPaids: paid,
      }),
    [options.projectPaidRelationships, therapeuticArea, indication, paid],
  );

  // Cascade for PAID options — filter by TA, Asset, Indication (NOT PAID)
  const cascadeForPaids = React.useMemo(
    () =>
      getCascadedOptionsMulti(options.projectPaidRelationships, {
        selectedTAs: therapeuticArea,
        selectedAssets: asset,
        selectedIndications: indication,
      }),
    [options.projectPaidRelationships, therapeuticArea, asset, indication],
  );

  // Sibling-safe Indication cascade: TA/Asset only.
  // Selecting a PAID should not auto-remove already selected indications.
  const cascadeForIndicationsStructural = React.useMemo(
    () =>
      getCascadedOptionsMulti(options.projectPaidRelationships, {
        selectedTAs: therapeuticArea,
        selectedAssets: asset,
      }),
    [options.projectPaidRelationships, therapeuticArea, asset],
  );

  // Sibling-safe PAID cascade: TA/Asset only.
  // Selecting an indication should not auto-remove already selected PAIDs.
  const cascadeForPaidsStructural = React.useMemo(
    () =>
      getCascadedOptionsMulti(options.projectPaidRelationships, {
        selectedTAs: therapeuticArea,
        selectedAssets: asset,
      }),
    [options.projectPaidRelationships, therapeuticArea, asset],
  );

  const focusAfterTopError = React.useCallback(() => {
    window.setTimeout(() => {
      focusTopError();
    }, 180);
  }, [focusTopError]);

  // ── Auto-clear invalid child selections when parent options change ──
  // Prune Sub-TA if current selection is no longer valid for selected TA
  React.useEffect(() => {
    if (fieldVisibility.subTherapeuticArea && subTherapeuticArea.length > 0) {
      const invalidSubTAs = subTherapeuticArea.filter(
        (value) => !filteredSubTherapeuticAreas.includes(value),
      );
      if (invalidSubTAs.length > 0) {
        setSubTherapeuticArea([]);
      }
    }
  }, [fieldVisibility.subTherapeuticArea, subTherapeuticArea, filteredSubTherapeuticAreas]);

  // Prune Disease Area if current selection is no longer valid for selected TA
  React.useEffect(() => {
    if (fieldVisibility.diseaseArea && diseaseArea.length > 0) {
      const invalidDAs = diseaseArea.filter(
        (value) => !filteredDiseaseAreaStrategies.includes(value),
      );
      if (invalidDAs.length > 0) {
        setDiseaseArea([]);
      }
    }
  }, [fieldVisibility.diseaseArea, diseaseArea, filteredDiseaseAreaStrategies]);

  // Prune Indication if current selection is no longer valid for cascade (TA/Asset/PAID)
  React.useEffect(() => {
    if (fieldVisibility.indication && indication.length > 0) {
      const validIndications = keepCascadeOverlap(
        indication,
        cascadeForIndicationsStructural.availableIndications,
      );
      if (validIndications.length !== indication.length) {
        // Preserve sibling behavior: keep overlap and avoid wiping all selections.
        if (validIndications.length > 0) {
          setIndication(validIndications);
        }
      }
    }
  }, [fieldVisibility.indication, indication, cascadeForIndicationsStructural.availableIndications]);

  // Prune Asset if current selection is no longer valid for cascade (TA/Indication/PAID)
  React.useEffect(() => {
    if (fieldVisibility.asset && asset.length > 0) {
      const invalidAssets = asset.filter(
        (value) => !cascadeForAssets.availableAssets.includes(value),
      );
      if (invalidAssets.length > 0) {
        setAsset([]);
      }
    }
  }, [fieldVisibility.asset, asset, cascadeForAssets.availableAssets]);

  // Prune PAID only for structural constraints (TA/Asset), not indication sibling narrowing.
  React.useEffect(() => {
    if (fieldVisibility.paid && paid.length > 0) {
      const validPaids = keepCascadeOverlap(
        paid,
        cascadeForPaidsStructural.availablePaids,
      );
      if (validPaids.length !== paid.length) {
        // Preserve sibling behavior: keep overlap and avoid wiping all selections.
        if (validPaids.length > 0) {
          setPaid(validPaids);
        }
      }
    }
  }, [fieldVisibility.paid, paid, cascadeForPaidsStructural.availablePaids]);

  // ── Auto-populate Sub-TA and LoT from ProjectPAID relationships ──
  // This effect triggers when Indication changes, computes available Sub-TA/LoT,
  // and auto-selects if exactly one valid option exists for a visible field
  React.useEffect(() => {
    // Extract unique Sub-TA values from relationships filtered by current selections
    const availableSubTAs = cascadeForPaids.availableSubTherapeuticAreas;
    const availableLoTs = cascadeForPaids.availableLineOfTherapies;

    // Auto-select Sub-TA if exactly one exists and field is visible
    if (
      fieldVisibility.subTherapeuticArea &&
      availableSubTAs.length === 1 &&
      subTherapeuticArea.length === 0
    ) {
      setSubTherapeuticArea([availableSubTAs[0]]);
    }

    // Auto-select LoT if exactly one exists and field is visible
    if (
      fieldVisibility.lineOfTherapy &&
      availableLoTs.length === 1 &&
      lineOfTherapy.length === 0
    ) {
      setLineOfTherapy([availableLoTs[0]]);
    }
  }, [indication, cascadeForPaids.availableSubTherapeuticAreas, cascadeForPaids.availableLineOfTherapies, fieldVisibility.subTherapeuticArea, fieldVisibility.lineOfTherapy, subTherapeuticArea, lineOfTherapy]);

  React.useEffect(() => {
    if (Object.keys(cascadeErrors).length === 0) return;
    setCascadeErrors({});
  }, [
    cascadeErrors,
    therapeuticArea,
    subTherapeuticArea,
    diseaseArea,
    indication,
    asset,
    paid,
  ]);

  const validateCascadeSelections = React.useCallback((): boolean => {
    const nextErrors: ICascadeFieldErrors = {};
    const buildInvalidSelectionMessage = (
      fieldLabel: string,
      invalidValues: string[],
      dependencyFields: string[],
    ): string => {
      const values = invalidValues.join(", ");
      return `${fieldLabel} has invalid selection(s): ${values}. Please update ${fieldLabel} to match the selected ${formatFieldLabels(dependencyFields)}.`;
    };

    const invalidSubTherapeuticAreas = subTherapeuticArea.filter(
      (value) => !filteredSubTherapeuticAreas.includes(value),
    );
    if (
      fieldVisibility.subTherapeuticArea &&
      invalidSubTherapeuticAreas.length > 0
    ) {
      nextErrors.subTherapeuticArea = buildInvalidSelectionMessage(
        "Sub-Therapeutic Area",
        invalidSubTherapeuticAreas,
        ["Therapeutic Area"],
      );
    }

    const invalidDiseaseAreas = diseaseArea.filter(
      (value) => !filteredDiseaseAreaStrategies.includes(value),
    );
    if (fieldVisibility.diseaseArea && invalidDiseaseAreas.length > 0) {
      nextErrors.diseaseArea = buildInvalidSelectionMessage(
        "Disease Area",
        invalidDiseaseAreas,
        ["Therapeutic Area"],
      );
    }

    const invalidIndications = indication.filter(
      (value) =>
        !containsCascadeValue(
          cascadeForIndicationsStructural.availableIndications,
          value,
        ),
    );
    if (fieldVisibility.indication && invalidIndications.length > 0) {
      nextErrors.indication = buildInvalidSelectionMessage(
        "Indication",
        invalidIndications,
        ["Therapeutic Area", "Asset"],
      );
    }

    const invalidAssets = asset.filter(
      (value) => !cascadeForAssets.availableAssets.includes(value),
    );
    if (fieldVisibility.asset && invalidAssets.length > 0) {
      nextErrors.asset = buildInvalidSelectionMessage("Asset", invalidAssets, [
        "Therapeutic Area",
        "Indication",
        "PAID",
      ]);
    }

    const invalidPaids = paid.filter(
      (value) =>
        !containsCascadeValue(cascadeForPaidsStructural.availablePaids, value),
    );
    if (fieldVisibility.paid && invalidPaids.length > 0) {
      nextErrors.paid = buildInvalidSelectionMessage("PAID", invalidPaids, [
        "Therapeutic Area",
        "Asset",
      ]);
    }

    setCascadeErrors(nextErrors);
    return Object.keys(nextErrors).length === 0;
  }, [
    fieldVisibility,
    subTherapeuticArea,
    filteredSubTherapeuticAreas,
    diseaseArea,
    filteredDiseaseAreaStrategies,
    indication,
    asset,
    paid,
    cascadeForIndicationsStructural.availableIndications,
    cascadeForAssets.availableAssets,
    cascadeForPaidsStructural.availablePaids,
  ]);

  const getMissingRequiredFieldLabels = React.useCallback((): string[] => {
    const missingFields: string[] = [];

    if (fieldRequirements.documentType && documentType.trim() === "") {
      missingFields.push("Document Type");
    }
    if (
      fieldVisibility.therapeuticArea &&
      fieldRequirements.therapeuticArea &&
      therapeuticArea.length === 0
    ) {
      missingFields.push("Therapeutic Area");
    }
    if (
      fieldVisibility.subTherapeuticArea &&
      fieldRequirements.subTherapeuticArea &&
      subTherapeuticArea.length === 0
    ) {
      missingFields.push("Sub-Therapeutic Area");
    }
    if (
      fieldVisibility.diseaseArea &&
      fieldRequirements.diseaseArea &&
      diseaseArea.length === 0
    ) {
      missingFields.push("Disease Area");
    }
    if (
      fieldVisibility.indication &&
      fieldRequirements.indication &&
      indication.length === 0
    ) {
      missingFields.push("Indication");
    }
    if (
      fieldVisibility.asset &&
      fieldRequirements.asset &&
      asset.length === 0
    ) {
      missingFields.push("Asset");
    }
    if (fieldVisibility.paid && fieldRequirements.paid && paid.length === 0) {
      missingFields.push("PAID");
    }
    if (
      fieldVisibility.lineOfTherapy &&
      fieldRequirements.lineOfTherapy &&
      lineOfTherapy.length === 0
    ) {
      missingFields.push("Line of Therapy");
    }
    if (
      fieldVisibility.effectiveDate &&
      fieldRequirements.effectiveDate &&
      effectiveDate.trim() === ""
    ) {
      missingFields.push("Effective Date");
    }
    if (
      fieldVisibility.comments &&
      fieldRequirements.comments &&
      comments.trim() === ""
    ) {
      missingFields.push("Comments");
    }

    return missingFields;
  }, [
    fieldRequirements,
    fieldVisibility,
    documentType,
    therapeuticArea,
    subTherapeuticArea,
    diseaseArea,
    indication,
    asset,
    paid,
    lineOfTherapy,
    effectiveDate,
    comments,
  ]);

  /**
   * Validate that field cardinality (single vs. multi-select) matches configuration.
   * Returns array of error messages for fields that violate cardinality constraints.
   */
  const getCardinalityViolations = React.useCallback((): string[] => {
    const violations: string[] = [];

    const fieldChecks: Array<{
      label: string;
      values: string[];
      fieldName: string;
    }> = [
      { label: "Therapeutic Area", values: therapeuticArea, fieldName: FIELD_NAMES.THERAPEUTIC_AREA },
      { label: "Sub-Therapeutic Area", values: subTherapeuticArea, fieldName: FIELD_NAMES.SUB_THERAPEUTIC_AREA },
      { label: "Asset", values: asset, fieldName: FIELD_NAMES.ASSET },
      { label: "Indication", values: indication, fieldName: FIELD_NAMES.INDICATION },
      { label: "Line of Therapy", values: lineOfTherapy, fieldName: FIELD_NAMES.LINE_OF_THERAPY },
      { label: "PAID", values: paid, fieldName: FIELD_NAMES.PAID },
      { label: "Disease Area", values: diseaseArea, fieldName: FIELD_NAMES.DISEASE_AREA },
    ];

    fieldChecks.forEach(({ label, values, fieldName }) => {
      if (!isFieldVisible(fieldName, documentType, selectedTherapeuticArea)) return;
      if (isFieldMultiSelect(fieldName, documentType, selectedTherapeuticArea)) return;
      if (values.length > 1) {
        violations.push(
          `${label} is single-select for ${documentType}, but ${values.length} values were selected. Please select only one ${label}.`,
        );
      }
    });

    return violations;
  }, [
    documentType,
    selectedTherapeuticArea,
    therapeuticArea,
    subTherapeuticArea,
    diseaseArea,
    indication,
    asset,
    paid,
    lineOfTherapy,
  ]);

  // ── Save handler ──────────────────────────────────────────
  const handleSave = React.useCallback(() => {
    if (!doc) return;

    setCascadeErrors({});

    const missingRequiredFields = getMissingRequiredFieldLabels();
    if (missingRequiredFields.length > 0) {
      setSaveError(
        `Missing required metadata: ${missingRequiredFields.join(", ")}.`,
      );
      focusAfterTopError();
      return;
    }

    // Validate field cardinality (single vs. multi-select)
    const cardinalityViolations = getCardinalityViolations();
    if (cardinalityViolations.length > 0) {
      setSaveError(
        `Invalid field selections: ${cardinalityViolations.join(" ")}`,
      );
      focusAfterTopError();
      return;
    }

    if (effectiveDate) {
      const parsedEffectiveDate = parseISO(effectiveDate);
      if (
        !isValid(parsedEffectiveDate) ||
        isFutureDate(parsedEffectiveDate)
      ) {
        setEffectiveDateError(FUTURE_EFFECTIVE_DATE_ERROR);
        setSaveError(FUTURE_EFFECTIVE_DATE_ERROR);
        focusAfterTopError();
        return;
      }
    }

    if (!validateCascadeSelections()) {
      setSaveError(INVALID_CASCADE_SELECTION_ERROR);
      focusAfterTopError();
      return;
    }

    setSaveError("");

    const searchTokens = buildDocumentSearchTokens(
      documentType ? [documentType] : [],
      asset,
      therapeuticArea,
      indication,
      options.documentTypes,
      options.assets,
      options.therapeuticAreas,
      options.indications,
    );

    onSave(doc.id, {
      documentType,
      therapeuticArea,
      subTherapeuticArea,
      asset,
      indication,
      lineOfTherapy,
      paid,
      diseaseArea,
      effectiveDate,
      comments,
      searchTokens,
    });
  }, [
    doc,
    documentType,
    therapeuticArea,
    subTherapeuticArea,
    asset,
    indication,
    lineOfTherapy,
    paid,
    diseaseArea,
    effectiveDate,
    comments,
    validateCascadeSelections,
    getCardinalityViolations,
    options,
    onSave,
    focusAfterTopError,
    getMissingRequiredFieldLabels,
  ]);

  // ── Helpers ───────────────────────────────────────────────
  const toOpts = (values: string[]): IDropdownOption[] =>
    values.map((v) => ({ key: v, text: v }));

  // ── Calendar day styles for blue selected/today state ────
  const calendarDayStyles = {
    daySelected: {
      backgroundColor: "#0066F5",
      color: "#ffffff",
      selectors: {
        "&:hover": {
          backgroundColor: "#0052CC",
          color: "#ffffff",
        },
        "& button": {
          color: "#ffffff",
        },
      },
    },
    dayIsToday: {
      backgroundColor: "#0066F5",
      color: "#ffffff",
      selectors: {
        "& button": {
          color: "#ffffff",
        },
      },
    },
  };

  if (!isOpen || !doc || !isAdmin) return null;

  return (
    <>
      <div className={styles.panelOverlay} onClick={onCancel} />
      <div
        className={`${styles.panel} ${doc.status === "Current" ? styles.panelCurrent : styles.panelArchived}`}
      >
        <div className={styles.panelToggleBar}>
          <button
            className={styles.panelToggleBtn}
            onClick={onCancel}
            aria-label="Close panel"
          >
            <img
              src={require("../../assets/icons/panel-toggle.svg")}
              alt=""
              style={{ width: "20px", height: "20px" }}
              aria-hidden="true"
            />
          </button>
          <hr className={styles.panelToggleDivider} />
        </div>
        <div className={styles.panelHeader}>
          <div className={styles.editPanelHeaderLead}>
            <span className={styles.panelTitle}>Edit</span>
            <span
              className={`${styles.editStatusPill} ${doc.status === "Current" ? styles.editStatusCurrent : styles.editStatusArchived}`}
            >
              {doc.status}
            </span>
          </div>
        </div>

        <div className={styles.panelBody} ref={panelBodyRef}>
          {(saveError || constraintError) && (
            <div
              className={styles.panelTopErrorBanner}
              ref={topErrorRef}
              tabIndex={-1}
              role="alert"
              aria-live="assertive"
            >
              {saveError && (
                <ErrorBanner
                  message={saveError}
                  onDismiss={() => setSaveError("")}
                />
              )}

              {constraintError && (
                <ErrorBanner
                  message={constraintError}
                  onDismiss={() => setConstraintError("")}
                />
              )}
            </div>
          )}

          {/* File name */}
          <div style={{ fontSize: 13, color: "#636363", marginBottom: 12 }}>
            <span className={styles.cellLink}>{doc.fileName}</span>
          </div>

          {/* Quick actions */}
          <div className={styles.quickActions}>
            <TooltipHost content="Replace File">
              <button
                className={styles.quickActionButton}
                onClick={onReplaceFileClick}
                aria-label="Replace file"
              >
                <img
                  src={require("../../assets/icons/replace.svg")}
                  alt=""
                  style={{
                    width: "1em",
                    height: "1em",
                    display: "inline-block",
                  }}
                  aria-hidden="true"
                />
                <span>Replace File</span>
              </button>
            </TooltipHost>

            {/* Archive (Current docs) or Re-Activate (Archived docs) */}
            {doc.status === "Current" ? (
              <TooltipHost content="Archive">
                <button
                  className={styles.quickActionButton}
                  onClick={onArchiveClick}
                  aria-label="Archive document"
                >
                  <img
                    src={require("../../assets/icons/archive.svg")}
                    alt=""
                    style={{
                      width: "16px",
                      height: "16px",
                      display: "inline-block",
                    }}
                    aria-hidden="true"
                  />
                  <span>Archive</span>
                </button>
              </TooltipHost>
            ) : (
              <TooltipHost content="Re-Activate">
                <button
                  className={styles.quickActionButton}
                  onClick={onReActivateClick}
                  aria-label="Re-activate document"
                >
                  <img
                    src={require("../../assets/icons/re-activate.svg")}
                    alt=""
                    style={{
                      width: "1em",
                      height: "1em",
                      display: "inline-block",
                    }}
                    aria-hidden="true"
                  />
                  <span>Re-Activate</span>
                </button>
              </TooltipHost>
            )}

            <TooltipHost content="Delete Document">
              <button
                className={`${styles.quickActionButton} ${styles.quickActionDanger}`}
                onClick={onDeleteClick}
                aria-label="Delete document"
              >
                <img
                  src={require("../../assets/icons/delete.svg")}
                  alt=""
                  style={{
                    width: "1em",
                    height: "1em",
                    display: "inline-block",
                  }}
                  aria-hidden="true"
                />
                <span>Delete</span>
              </button>
            </TooltipHost>
          </div>

          {/* 1. Document Type */}
          <div className={styles.formGroup}>
            <Dropdown
              label="Document Type"
              ariaLabel="Document Type"
              required
              options={toOpts(toValueArray(options.documentTypes))}
              selectedKey={documentType || undefined}
              onChange={(_, item) =>
                setDocumentType(item ? (item.key as string) : "")
              }
              placeholder="Select document type"
            />
          </div>

          {/* 2. Therapeutic Area */}
          {fieldVisibility.therapeuticArea && (
            <div className={styles.formGroup}>
              <Dropdown
                label="Therapeutic Area"
                ariaLabel="Therapeutic Area"
                required={fieldRequirements.therapeuticArea}
                options={toOpts(toValueArray(options.therapeuticAreas))}
                selectedKey={
                  therapeuticArea.length > 0 ? therapeuticArea[0] : undefined
                }
                onChange={(_, item) =>
                  setTherapeuticArea(item ? [item.key as string] : [])
                }
                placeholder="Select therapeutic area"
              />
            </div>
          )}

          {/* 3. Sub-Therapeutic Area */}
          {fieldVisibility.subTherapeuticArea && (
            <SearchableDropdown
              label="Sub-Therapeutic Area"
              required={fieldRequirements.subTherapeuticArea}
              options={filteredSubTherapeuticAreas}
              selectedKeys={subTherapeuticArea}
              onChange={setSubTherapeuticArea}
              errorMessage={cascadeErrors.subTherapeuticArea}
              placeholder="Select sub-therapeutic area"
              multiSelect={false}
              showChipsBelow={true}
            />
          )}

          {/* 4. Disease Area Strategy */}
          {fieldVisibility.diseaseArea && (
            <SearchableDropdown
              label="Disease Area"
              required={fieldRequirements.diseaseArea}
              options={filteredDiseaseAreaStrategies}
              selectedKeys={diseaseArea}
              onChange={setDiseaseArea}
              errorMessage={cascadeErrors.diseaseArea}
              placeholder="Type to search disease area..."
              multiSelect={fieldMultiSelect.diseaseArea}
              showChipsBelow={true}
            />
          )}

          {/* 5. Asset */}
          {fieldVisibility.asset && (
            <SearchableDropdown
              label="Asset"
              required={fieldRequirements.asset}
              options={cascadeForAssets.availableAssets}
              selectedKeys={asset}
              onChange={setAsset}
              errorMessage={cascadeErrors.asset}
              placeholder="Select asset"
              multiSelect={false}
              showChipsBelow={true}
            />
          )}

          {/* 6. Indication */}
          {fieldVisibility.indication && (
            <SearchableDropdown
              label="Indication"
              required={fieldRequirements.indication}
              options={
                cascadeForIndicationsStructural.availableIndications.length > 0
                  ? cascadeForIndicationsStructural.availableIndications
                  : toValueArray(options.indications)
              }
              selectedKeys={indication}
              onChange={setIndication}
              errorMessage={cascadeErrors.indication}
              placeholder="Select indication"
              multiSelect={fieldMultiSelect.indication}
              showChipsBelow={true}
            />
          )}

          {/* 7. Line of Therapy */}
          {fieldVisibility.lineOfTherapy && (
            <SearchableDropdown
              label="Line of Therapy"
              required={fieldRequirements.lineOfTherapy}
              options={toValueArray(options.lineOfTherapy)}
              selectedKeys={lineOfTherapy}
              onChange={setLineOfTherapy}
              placeholder="Type to search line of therapy..."
              multiSelect={true}
              showChipsBelow={true}
            />
          )}

          {/* 8. PAID */}
          {fieldVisibility.paid && (
            <SearchableDropdown
              label="PAID"
              required={fieldRequirements.paid}
              options={cascadeForPaidsStructural.availablePaids}
              selectedKeys={paid}
              onChange={setPaid}
              errorMessage={cascadeErrors.paid}
              placeholder="Select PAID"
              multiSelect={fieldMultiSelect.paid}
              showChipsBelow={true}
            />
          )}

          {/* 9. Effective Date */}
          {fieldVisibility.effectiveDate && (
            <div>
              <div className={styles.formGroup}>
                <DatePicker
                  label="Effective Date"
                  isRequired={fieldRequirements.effectiveDate}
                  allowTextInput={true}
                  maxDate={new Date()}
                  calendarProps={{
                    calendarDayProps: {
                      styles: calendarDayStyles,
                    },
                  }}
                  value={toPickerDate(effectiveDate)}
                  onSelectDate={(date) => {
                    if (!date) {
                      setEffectiveDate("");
                      setEffectiveDateError("");
                      setSaveError("");
                      return;
                    }

                    if (isFutureDate(date)) {
                      setEffectiveDateError(FUTURE_EFFECTIVE_DATE_ERROR);
                      return;
                    }

                    setEffectiveDate(format(date, "yyyy-MM-dd"));
                    setEffectiveDateError("");
                    setSaveError("");
                  }}
                  formatDate={(date) => {
                    return date ? format(date, "MM/dd/yyyy") : "";
                  }}
                  parseDateFromString={(dateStr) => {
                    const parsed = parseDateFns(
                      dateStr,
                      "MM/dd/yyyy",
                      new Date(),
                    );
                    return isValid(parsed) ? parsed : null;
                  }}
                  textField={{
                    errorMessage: effectiveDateError,
                    ariaLabel: "Effective Date",
                  }}
                  placeholder="Select date"
                />
                <div style={{ fontSize: "12px", color: "#6b7280", marginTop: "4px" }}>
                  The date the document was aligned at TASC/AST.
                </div>
              </div>
            </div>
          )}

          {/* 10. Comments */}
          {fieldVisibility.comments && (
            <div className={styles.formGroup}>
              <TextField
                label="Comments"
                ariaLabel="Comments"
                required={fieldRequirements.comments}
                multiline
                rows={3}
                value={comments}
                onChange={(_, val) =>
                  setComments((val || "").slice(0, COMMENTS_MAX_LENGTH))
                }
                placeholder="Enter comments"
                maxLength={COMMENTS_MAX_LENGTH}
              />
              <div className={styles.formCharCount}>
                {comments.length}/{COMMENTS_MAX_LENGTH}
              </div>
            </div>
          )}

          {/* Uploaded By (511) — read-only system field */}
          <div className={styles.formGroup}>
            <TextField
              label="Uploaded By (511)"
              value={doc.spectra511}
              readOnly
              disabled
            />
          </div>

          <div className={styles.formNotice}>
            Once a document is submitted, only an Admin can edit, archive or
            delete. Please check that the correct selections are made.
          </div>
        </div>

        <div className={styles.panelFooter}>
          <button className={styles.btnSecondary} onClick={onCancel}>
            Cancel
          </button>
          <button
            className={styles.btnPrimary}
            onClick={handleSave}
            disabled={isSaving}
          >
            {isSaving ? (
              <>
                <span className={styles.spinner} /> Saving...
              </>
            ) : (
              <>
                <img
                  src={require("../../assets/icons/save.svg")}
                  alt=""
                  style={{
                    width: "16px",
                    height: "16px",
                    display: "inline-block",
                    marginRight: "6px",
                  }}
                  aria-hidden="true"
                />
                Save Updates
              </>
            )}
          </button>
        </div>
      </div>
    </>
  );
};
