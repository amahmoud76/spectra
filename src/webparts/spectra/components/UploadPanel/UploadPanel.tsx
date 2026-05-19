import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import {
  IMetadataOptions,
  toValueArray,
} from "../../interfaces/IMetadataOptions";
import { IUploadPayload } from "../../interfaces/IUploadPayload";
import {
  getCascadedOptionsMulti,
  getAllPaids,
  getDiseaseAreaStrategiesForTherapeuticArea,
  // ICascadeSelection,
} from "../../utils/cascadingFilterHelper";
import { buildDocumentSearchTokens } from "../../utils/searchTokenHelper";
import { validateFile } from "../../utils/fileHelper";
import {
  ADMIN_ONLY_DOC_TYPES,
  COMMENTS_MAX_LENGTH,
  ACCEPTED_FILE_DISPLAY,
} from "../../config/config";
import {
  FIELD_NAMES,
  isFieldVisible,
  isFieldRequired,
  isFieldMultiSelect,
} from "../../config/fieldConfig";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { DatePicker } from "@fluentui/react/lib/DatePicker";
import { TextField } from "@fluentui/react/lib/TextField";
import { SearchableDropdown } from "../SearchableDropdown/SearchableDropdown";
import { ErrorBanner } from "../ErrorBanner/ErrorBanner";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { generateFileName } from "../../utils/fileNamingHelper";
import {
  format,
  isValid as isValidDate,
  parse as parseDateFns,
  parseISO,
} from "date-fns";
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
  return isValidDate(parsed) ? parsed : undefined;
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

const CONTRIBUTOR_EDITABLE_DOC_TYPES = new Set<string>([
  "R&D One-Pager",
  "DAS",
  "PDS",
]);

export interface IUploadPanelProps {
  isOpen: boolean;
  options: IMetadataOptions;
  isAdmin: boolean;
  isContributor?: boolean;
  isUploading: boolean;
  uploadProgress?:
    | "Preparing"
    | "UploadingFile"
    | "SavingMetadata"
    | "Finalizing"
    | "Completed"
    | "Failed"
    | "";
  uploadPercent?: number;
  cancelUpload?: () => void;
  panelNotice?: string;
  onSubmit: (payload: IUploadPayload) => void;
  onCancel: () => void;
  title?: string;
  prePopulatedData?: Partial<IUploadPayload>;
  archiveTargetDocument?: IDocument;
}

const formatFileSize = (bytes: number): string => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

export const UploadPanel: React.FC<IUploadPanelProps> = ({
  isOpen,
  options,
  isAdmin,
  isContributor = false,
  isUploading,
  uploadProgress = "",
  uploadPercent = 0,
  cancelUpload,
  panelNotice,
  onSubmit,
  onCancel,
  title = "Upload Document",
  prePopulatedData,
  archiveTargetDocument,
}) => {
  const prevIsOpenRef = React.useRef(false);

  // ── Form state ────────────────────────────────────────────
  const [file, setFile] = React.useState<File | null>(null);
  const [fileError, setFileError] = React.useState("");
  const [documentType, setDocumentType] = React.useState(
    prePopulatedData?.documentType || "",
  );
  const [therapeuticArea, setTherapeuticArea] = React.useState(
    prePopulatedData?.therapeuticArea || ([] as string[]),
  );
  const [subTherapeuticArea, setSubTherapeuticArea] = React.useState(
    prePopulatedData?.subTherapeuticArea || ([] as string[]),
  );
  const [asset, setAsset] = React.useState(
    prePopulatedData?.asset || ([] as string[]),
  );
  const [indication, setIndication] = React.useState(
    prePopulatedData?.indication || ([] as string[]),
  );
  const [lineOfTherapy, setLineOfTherapy] = React.useState(
    prePopulatedData?.lineOfTherapy || ([] as string[]),
  );
  const [paid, setPaid] = React.useState(
    prePopulatedData?.paid || ([] as string[]),
  );
  const [diseaseArea, setDiseaseArea] = React.useState<string[]>(
    prePopulatedData?.diseaseArea || [],
  );
  const [effectiveDate, setEffectiveDate] = React.useState(
    prePopulatedData?.effectiveDate || "",
  );
  const [description, setDescription] = React.useState(
    prePopulatedData?.description || "",
  );
  const [comments, setComments] = React.useState(
    prePopulatedData?.comments || "",
  );
  const [effectiveDateError, setEffectiveDateError] = React.useState("");
  const [cascadeErrors, setCascadeErrors] = React.useState<ICascadeFieldErrors>(
    {},
  );
  const [submitError, setSubmitError] = React.useState("");
  const [constraintError, setConstraintError] = React.useState("");

  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const topErrorRef = React.useRef<HTMLDivElement>(null);
  const panelBodyRef = React.useRef<HTMLDivElement>(null);

  const focusTopError = React.useCallback(() => {
    if (!topErrorRef.current) return;
    topErrorRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
    topErrorRef.current.focus();
  }, []);

  React.useEffect(() => {
    if (!submitError && !constraintError) return;
    // Wait for the banner to mount before scrolling/focusing.
    requestAnimationFrame(() => {
      focusTopError();
    });
  }, [submitError, constraintError, focusTopError]);

  const availableDocTypes = React.useMemo(
    () =>
      toValueArray(options.documentTypes).filter(
        (dt) => isAdmin || !ADMIN_ONLY_DOC_TYPES.includes(dt),
      ),
    [options.documentTypes, isAdmin],
  );
  const resetFormState = React.useCallback((seed?: Partial<IUploadPayload>) => {
    setFile(null);
    setFileError("");
    setDocumentType(seed?.documentType || "");
    setTherapeuticArea(seed?.therapeuticArea ? [...seed.therapeuticArea] : []);
    setSubTherapeuticArea(
      seed?.subTherapeuticArea ? [...seed.subTherapeuticArea] : [],
    );
    setAsset(seed?.asset ? [...seed.asset] : []);
    setIndication(seed?.indication ? [...seed.indication] : []);
    setLineOfTherapy(seed?.lineOfTherapy ? [...seed.lineOfTherapy] : []);
    setPaid(seed?.paid ? [...seed.paid] : []);
    setDiseaseArea(seed?.diseaseArea ? [...seed.diseaseArea] : []);
    setEffectiveDate(seed?.effectiveDate || "");
    setDescription(seed?.description || "");
    setComments(seed?.comments || "");
    setEffectiveDateError("");
    setCascadeErrors({});
    setSubmitError("");

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }, []);

  React.useEffect(() => {
    if (isOpen && !prevIsOpenRef.current) {
      resetFormState(prePopulatedData);
    }

    prevIsOpenRef.current = isOpen;
  }, [isOpen, prePopulatedData, resetFormState]);

  // ── Derived field state — visibility, requirements, multi-select ──
  const selectedTherapeuticArea =
    therapeuticArea.length > 0 ? therapeuticArea[0] : "";

  // In Archive & Replace mode, contributors can only edit TA/Asset for exempt doc types.
  const shouldLockContributorTaAndAsset =
    Boolean(archiveTargetDocument) &&
    isContributor &&
    !CONTRIBUTOR_EDITABLE_DOC_TYPES.has(documentType);

  // Contributors cannot change Document Type during Archive & Replace.
  const shouldLockContributorDocType =
    Boolean(archiveTargetDocument) && isContributor;

  const fieldVisibility = React.useMemo(
    () => ({
      documentType: isFieldVisible(
        FIELD_NAMES.DOCUMENT_TYPE,
        documentType,
        selectedTherapeuticArea,
      ),
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
      description: isFieldVisible(
        FIELD_NAMES.DESCRIPTION,
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
      description: isFieldRequired(
        FIELD_NAMES.DESCRIPTION,
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

  // ── Cascading — separate cascade per field ────────────────
  // Each field's cascade EXCLUDES its own selection to prevent
  // self-filtering (which would collapse the options list to
  // only the already-selected value).

  // Cascade for TA options — filter by Asset, Indication, PAID (NOT TA)
  const cascadeForTAs = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedAssets: asset,
      selectedIndications: indication,
      selectedPaids: paid,
    },
  );

  // Cascade for Asset options — filter by TA, Indication, PAID (NOT Asset)
  const cascadeForAssets = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: therapeuticArea,
      selectedIndications: indication,
      selectedPaids: paid,
    },
  );

  // Cascade for PAID options — filter by TA, Asset, Indication (NOT PAID)
  const cascadeForPaids = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: therapeuticArea,
      selectedAssets: asset,
      selectedIndications: indication,
    },
  );

  // Sibling-safe Indication cascade: TA/Asset only.
  // Selecting a PAID should not auto-remove already selected indications.
  const cascadeForIndicationsStructural = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: therapeuticArea,
      selectedAssets: asset,
    },
  );

  // Sibling-safe PAID cascade: TA/Asset only.
  // Selecting an indication should not auto-remove already selected PAIDs.
  const cascadeForPaidsStructural = getCascadedOptionsMulti(
    options.projectPaidRelationships,
    {
      selectedTAs: therapeuticArea,
      selectedAssets: asset,
    },
  );

  const allPaids = getAllPaids(options.projectPaidRelationships);
  const allIndications = toValueArray(options.indications);
  const allAssets = toValueArray(options.assets);

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

  const filteredDiseaseAreaStrategies = React.useMemo(
    () =>
      getDiseaseAreaStrategiesForTherapeuticArea(
        options.diseaseAreaStrategyRelationships,
        selectedTherapeuticArea,
      ),
    [options.diseaseAreaStrategyRelationships, selectedTherapeuticArea],
  );

  const focusAfterTopError = React.useCallback(() => {
    window.setTimeout(() => {
      focusTopError();
    }, 180);
  }, [focusTopError]);

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

  // ── Clear hidden fields when they become hidden ──────────
  React.useEffect(() => {
    // Clear Sub-TA if it becomes hidden (TA != Aesthetics)
    if (!fieldVisibility.subTherapeuticArea && subTherapeuticArea.length > 0) {
      setSubTherapeuticArea([]);
    }
    // Clear Line of Therapy if it becomes hidden (TA != Oncology)
    if (!fieldVisibility.lineOfTherapy && lineOfTherapy.length > 0) {
      setLineOfTherapy([]);
    }
    // Clear Indication if it becomes hidden
    if (!fieldVisibility.indication && indication.length > 0) {
      setIndication([]);
    }
    // Clear Asset if it becomes hidden
    if (!fieldVisibility.asset && asset.length > 0) {
      setAsset([]);
    }
    // Clear PAID if it becomes hidden
    if (!fieldVisibility.paid && paid.length > 0) {
      setPaid([]);
    }
    // Clear Disease Area if it becomes hidden
    if (!fieldVisibility.diseaseArea && diseaseArea.length > 0) {
      setDiseaseArea([]);
    }
    // Clear Therapeutic Area if it becomes hidden
    if (!fieldVisibility.therapeuticArea && therapeuticArea.length > 0) {
      setTherapeuticArea([]);
    }
  }, [
    fieldVisibility,
    subTherapeuticArea,
    lineOfTherapy,
    indication,
    asset,
    paid,
    diseaseArea,
    therapeuticArea,
  ]);

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

  // ── File handling ─────────────────────────────────────────
  const handleFileSelect = React.useCallback((selectedFile: File) => {
    const validation = validateFile(selectedFile);
    if (!validation.isValid) {
      setFileError(validation.errorMessage);
      setFile(null);
      return;
    }
    setFileError("");
    setFile(selectedFile);
  }, []);

  const handleDrop = React.useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile) handleFileSelect(droppedFile);
    },
    [handleFileSelect],
  );

  const handleBrowse = React.useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileInputChange = React.useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const selectedFile = e.target.files?.[0];
      if (selectedFile) handleFileSelect(selectedFile);
    },
    [handleFileSelect],
  );

  const handleRemoveFile = React.useCallback(() => {
    setFile(null);
    setFileError("");
    setCascadeErrors({});
    setSubmitError("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  }, []);

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

    const validIndications =
      cascadeForIndicationsStructural.availableIndications.length > 0
        ? cascadeForIndicationsStructural.availableIndications
        : allIndications;
    const validAssets =
      cascadeForAssets.availableAssets.length > 0
        ? cascadeForAssets.availableAssets
        : allAssets;
    const validPaids =
      cascadeForPaidsStructural.availablePaids.length > 0
        ? cascadeForPaidsStructural.availablePaids
        : allPaids;

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
      (value) => !containsCascadeValue(validIndications, value),
    );
    if (fieldVisibility.indication && invalidIndications.length > 0) {
      nextErrors.indication = buildInvalidSelectionMessage(
        "Indication",
        invalidIndications,
        ["Therapeutic Area", "Asset"],
      );
    }

    const invalidAssets = asset.filter((value) => !validAssets.includes(value));
    if (fieldVisibility.asset && invalidAssets.length > 0) {
      nextErrors.asset = buildInvalidSelectionMessage("Asset", invalidAssets, [
        "Therapeutic Area",
        "Indication",
        "PAID",
      ]);
    }

    const invalidPaids = paid.filter(
      (value) => !containsCascadeValue(validPaids, value),
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
    allIndications,
    allAssets,
    allPaids,
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
      fieldVisibility.description &&
      fieldRequirements.description &&
      description.trim() === ""
    ) {
      missingFields.push("Description");
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
    description,
    comments,
  ]);

  /**
   * Validate that field cardinality (single vs. multi-select) matches configuration.
   * For example, EBP requires REQUIRED_SINGLE Indication, so multiple indications are invalid.
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

  // ── Validation ────────────────────────────────────────────
  const isValid = React.useMemo(() => {
    // File must always be present
    if (file === null || fileError) return false;

    // Document Type must always be set (required)
    if (documentType === "") return false;

    // Check required visible fields dynamically
    if (fieldRequirements.therapeuticArea && therapeuticArea.length === 0)
      return false;
    if (fieldRequirements.diseaseArea && diseaseArea.length === 0) return false;
    if (fieldRequirements.indication && indication.length === 0) return false;
    if (fieldRequirements.asset && asset.length === 0) return false;
    if (fieldRequirements.paid && paid.length === 0) return false;
    if (fieldRequirements.lineOfTherapy && lineOfTherapy.length === 0)
      return false;
    if (fieldRequirements.effectiveDate && effectiveDate === "") return false;
    if (effectiveDate) {
      const parsedEffectiveDate = parseISO(effectiveDate);
      if (
        !isValidDate(parsedEffectiveDate) ||
        isFutureDate(parsedEffectiveDate)
      ) {
        return false;
      }
    }
    if (fieldRequirements.comments && comments.trim() === "") return false;
    if (fieldRequirements.subTherapeuticArea && subTherapeuticArea.length === 0)
      return false;

    return true;
  }, [
    file,
    fileError,
    documentType,
    therapeuticArea,
    diseaseArea,
    indication,
    asset,
    paid,
    lineOfTherapy,
    effectiveDate,
    effectiveDateError,
    comments,
    subTherapeuticArea,
    fieldRequirements,
  ]);

  // ── Generate filename preview ────────────────────────────
  const generatedFileNamePreview = React.useMemo(() => {
    if (!file || !documentType) return "";
    try {
      const payload: IUploadPayload = {
        file,
        documentType,
        therapeuticArea,
        subTherapeuticArea,
        asset,
        indication,
        lineOfTherapy,
        paid,
        diseaseArea,
        effectiveDate,
        description,
        comments,
        searchTokens: [],
      };
      return generateFileName(payload, payload.file.name.split(".").pop() || "pdf");
    } catch {
      return "";
    }
  }, [
    file,
    documentType,
    therapeuticArea,
    subTherapeuticArea,
    asset,
    indication,
    lineOfTherapy,
    paid,
    diseaseArea,
    effectiveDate,
    description,
    comments,
  ]);

  // ── Submit ────────────────────────────────────────────────
  const handleSubmit = React.useCallback(() => {
    // Check for missing file on submit attempt
    if (!file) {
      setSubmitError("File is required. Please select a file to upload.");
      focusAfterTopError();
      return;
    }

    // Clear submit error if file is present
    setSubmitError("");
    setCascadeErrors({});

    const missingRequiredFields = getMissingRequiredFieldLabels();
    if (missingRequiredFields.length > 0) {
      setSubmitError(
        `Missing required metadata: ${missingRequiredFields.join(", ")}.`,
      );
      focusAfterTopError();
      return;
    }

    // Validate field cardinality (single vs. multi-select)
    const cardinalityViolations = getCardinalityViolations();
    if (cardinalityViolations.length > 0) {
      setSubmitError(
        `Invalid field selections: ${cardinalityViolations.join(" ")}`,
      );
      focusAfterTopError();
      return;
    }

    if (effectiveDate) {
      const parsedEffectiveDate = parseISO(effectiveDate);
      if (
        !isValidDate(parsedEffectiveDate) ||
        isFutureDate(parsedEffectiveDate)
      ) {
        setEffectiveDateError(FUTURE_EFFECTIVE_DATE_ERROR);
        setSubmitError(FUTURE_EFFECTIVE_DATE_ERROR);
        focusAfterTopError();
        return;
      }
    }

    if (!isValid) {
      setSubmitError("Please complete all required metadata fields.");
      focusAfterTopError();
      return;
    }

    if (!validateCascadeSelections()) {
      setSubmitError(INVALID_CASCADE_SELECTION_ERROR);
      focusAfterTopError();
      return;
    }

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

    const payload: IUploadPayload = {
      file,
      documentType,
      therapeuticArea,
      subTherapeuticArea,
      asset,
      indication,
      lineOfTherapy,
      paid,
      diseaseArea,
      effectiveDate,
      description,
      comments,
      searchTokens,
    };

    onSubmit(payload);
  }, [
    file,
    isValid,
    documentType,
    therapeuticArea,
    subTherapeuticArea,
    asset,
    indication,
    lineOfTherapy,
    paid,
    diseaseArea,
    effectiveDate,
    effectiveDateError,
    validateCascadeSelections,
    getCardinalityViolations,
    focusAfterTopError,
    description,
    comments,
    options,
    onSubmit,
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

  if (!isOpen) return null;

  return (
    <>
      <div className={styles.panelOverlay} onClick={onCancel} />
      <div className={styles.panel}>
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
          <span className={styles.panelTitle}>{title}</span>
        </div>

        <div className={styles.panelBody} ref={panelBodyRef}>
          {panelNotice && (
            <div className={styles.panelSuccessNotice} role="status">
              <img
                src={require("../../assets/icons/check.svg")}
                alt=""
                aria-hidden="true"
                style={{
                  width: "18px",
                  height: "18px",
                  display: "block",
                  flexShrink: 0,
                }}
              />
              <span>{panelNotice}</span>
            </div>
          )}

          {(submitError || constraintError) && (
            <div
              className={styles.panelTopErrorBanner}
              ref={topErrorRef}
              tabIndex={-1}
              role="alert"
              aria-live="assertive"
            >
              {/* ── Submit Error Banner ──────────────────────────── */}
              {submitError && (
                <ErrorBanner
                  message={submitError}
                  onDismiss={() => setSubmitError("")}
                />
              )}

              {/* ── Constraint Violation Banner ──────────────────── */}
              {constraintError && (
                <ErrorBanner
                  message={constraintError}
                  onDismiss={() => setConstraintError("")}
                />
              )}
            </div>
          )}

          {/* ── File Drop Zone ──────────────────────────────── */}
          <div
            className={styles.dropZone}
            onDrop={handleDrop}
            onDragOver={(e) => e.preventDefault()}
            onClick={handleBrowse}
            role="button"
            tabIndex={0}
            aria-label="Drop file here or click to browse"
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".pdf,.docx,.pptx,.xlsx"
              onChange={handleFileInputChange}
              style={{ display: "none" }}
            />
            <div className={styles.dropZoneIcon}>
              <img
                src={require("../../assets/icons/cloud.svg")}
                alt=""
                style={{ width: "48px", height: "48px" }}
                aria-hidden="true"
              />
            </div>
            <div className={styles.dropZoneText}>
              <strong>Drag & Drop</strong> or{" "}
              <span className={styles.dropZoneBrowse}>Browse Files</span>
            </div>
            <div className={styles.dropZoneNote}>
              File names will be standardized once uploaded.
              <br />
              Accepted Files: {ACCEPTED_FILE_DISPLAY}
            </div>
          </div>

          {file && (
            <div className={styles.dropZoneFile}>
              <span>
                {file.name}
                <span style={{ color: "var(--spectra-text-secondary)", marginLeft: 6, fontWeight: 400 }}>
                  ({formatFileSize(file.size)})
                </span>
              </span>
              <TooltipHost content="Remove file">
                <button
                  className={styles.dropZoneFileRemove}
                  onClick={handleRemoveFile}
                  aria-label="Remove file"
                >
                  <svg
                    width="12"
                    height="12"
                    viewBox="0 0 12 12"
                    fill="none"
                    aria-hidden="true"
                  >
                    <path
                      d="M1 1L11 11M11 1L1 11"
                      stroke="currentColor"
                      strokeWidth="1.75"
                      strokeLinecap="round"
                    />
                  </svg>
                </button>
              </TooltipHost>
            </div>
          )}
          {fileError && <div className={styles.formErrorText}>{fileError}</div>}

          {/* ── Filename Preview with Naming Convention Guidance ─── */}
          {file && documentType && (
            <div className={styles.formGroup}>
              <div className={styles.namingCard}>
                <div className={styles.namingHeader}>
                  <img
                    src={require("../../assets/icons/circle-exclamation.svg")}
                    alt=""
                    aria-hidden="true"
                    style={{
                      width: "16px",
                      height: "16px",
                      display: "block",
                      flexShrink: 0,
                    }}
                  />
                  <strong className={styles.namingTitle}>
                    File Naming Convention
                  </strong>
                </div>

                <div className={styles.namingText}>
                  Your file will be renamed to follow the naming convention when
                  uploaded:
                </div>

                <div className={styles.namingImmutablePreview}>
                  {generatedFileNamePreview}
                </div>

                <div className={styles.namingSubText}>
                  Original filename <strong>&quot;{file.name}&quot;</strong>{" "}
                  will be preserved for reference.
                </div>
              </div>
            </div>
          )}

          {/* ── Metadata Form — Updated field order ── */}

          {/* 1. Document Type — 9 items, keep as Dropdown */}
          <div className={styles.formGroup}>
            <Dropdown
              label="Document Type"
              ariaLabel="Document Type"
              required
              disabled={shouldLockContributorDocType}
              options={toOpts(availableDocTypes)}
              selectedKey={documentType || undefined}
              onChange={(_, item) =>
                setDocumentType(item ? (item.key as string) : "")
              }
              placeholder="Select document type"
            />
          </div>

          {
            <>
              {/* 2. Therapeutic Area — 6 items, Dropdown with cascade */}
              {fieldVisibility.therapeuticArea && (
                <div className={styles.formGroup}>
                  <Dropdown
                    label="Therapeutic Area"
                    ariaLabel="Therapeutic Area"
                    required={fieldRequirements.therapeuticArea}
                    disabled={shouldLockContributorTaAndAsset}
                    options={toOpts(
                      cascadeForTAs.availableTAs.length > 0
                        ? cascadeForTAs.availableTAs
                        : toValueArray(options.therapeuticAreas),
                    )}
                    selectedKey={
                      therapeuticArea.length > 0
                        ? therapeuticArea[0]
                        : undefined
                    }
                    onChange={(_, item) =>
                      setTherapeuticArea(item ? [item.key as string] : [])
                    }
                    placeholder={
                      fieldRequirements.therapeuticArea
                        ? "Select therapeutic area (required)"
                        : "Select therapeutic area (optional)"
                    }
                  />
                </div>
              )}

              {/* 3. Sub-Therapeutic Area — conditionally shown for Aesthetics TA */}
              {fieldVisibility.subTherapeuticArea && (
                <SearchableDropdown
                  label="Sub-Therapeutic Area"
                  required={fieldRequirements.subTherapeuticArea}
                  options={filteredSubTherapeuticAreas}
                  selectedKeys={subTherapeuticArea}
                  onChange={setSubTherapeuticArea}
                  errorMessage={cascadeErrors.subTherapeuticArea}
                  placeholder={
                    selectedTherapeuticArea
                      ? fieldRequirements.subTherapeuticArea
                        ? "Type to search sub-therapeutic area (required)..."
                        : "Type to search sub-therapeutic area (optional)..."
                      : "Select therapeutic area first"
                  }
                  multiSelect={false}
                  showChipsBelow={true}
                  disabled={!selectedTherapeuticArea}
                />
              )}

              {/* 4. Disease Area Strategy — cascaded by Therapeutic Area */}
              {fieldVisibility.diseaseArea && (
                <SearchableDropdown
                  label="Disease Area"
                  required={fieldRequirements.diseaseArea}
                  options={filteredDiseaseAreaStrategies}
                  selectedKeys={diseaseArea}
                  onChange={setDiseaseArea}
                  errorMessage={cascadeErrors.diseaseArea}
                  placeholder={
                    fieldRequirements.diseaseArea
                      ? "Type to search disease area (required)..."
                      : "Type to search disease area (optional)..."
                  }
                  multiSelect={fieldMultiSelect.diseaseArea}
                  showChipsBelow={true}
                />
              )}

              {/* 5. Asset — conditionally shown */}
              {fieldVisibility.asset && (
                <SearchableDropdown
                  label="Asset"
                  required={fieldRequirements.asset}
                  disabled={shouldLockContributorTaAndAsset}
                  options={
                    cascadeForAssets.availableAssets.length > 0
                      ? cascadeForAssets.availableAssets
                      : toValueArray(options.assets)
                  }
                  selectedKeys={asset}
                  onChange={setAsset}
                  errorMessage={cascadeErrors.asset}
                  placeholder={
                    fieldRequirements.asset
                      ? "Type to search asset (required)..."
                      : "Type to search asset (optional)..."
                  }
                  multiSelect={false}
                  showChipsBelow={true}
                />
              )}

              {/* 6. Indication — multi-select, cascade adjusts available options */}
              {fieldVisibility.indication && (
                <SearchableDropdown
                  label="Indication"
                  required={fieldRequirements.indication}
                  options={
                    cascadeForIndicationsStructural.availableIndications
                      .length > 0
                      ? cascadeForIndicationsStructural.availableIndications
                      : toValueArray(options.indications)
                  }
                  selectedKeys={indication}
                  onChange={setIndication}
                  errorMessage={cascadeErrors.indication}
                  placeholder={
                    fieldRequirements.indication
                      ? "Type to search indication (required)..."
                      : "Type to search indication (optional)..."
                  }
                  multiSelect={fieldMultiSelect.indication}
                  showChipsBelow={true}
                />
              )}

              {/* 7. Line of Therapy — conditionally shown for Oncology TA */}
              {fieldVisibility.lineOfTherapy && (
                <SearchableDropdown
                  label="Line of Therapy"
                  required={fieldRequirements.lineOfTherapy}
                  options={toValueArray(options.lineOfTherapy)}
                  selectedKeys={lineOfTherapy}
                  onChange={setLineOfTherapy}
                  placeholder={
                    fieldRequirements.lineOfTherapy
                      ? "Type to search line of therapy (required)..."
                      : "Type to search line of therapy (optional)..."
                  }
                  multiSelect={true}
                  showChipsBelow={true}
                />
              )}

              {/* 8. PAID — conditionally shown and multi-select per doc type */}
              {fieldVisibility.paid && (
                <SearchableDropdown
                  label="PAID"
                  required={fieldRequirements.paid}
                  options={
                    cascadeForPaidsStructural.availablePaids.length > 0
                      ? cascadeForPaidsStructural.availablePaids
                      : allPaids
                  }
                  selectedKeys={paid}
                  onChange={setPaid}
                  errorMessage={cascadeErrors.paid}
                  placeholder={
                    fieldRequirements.paid
                      ? "Type to search PAID (required)..."
                      : "Type to search PAID (optional)..."
                  }
                  multiSelect={fieldMultiSelect.paid}
                  showChipsBelow={true}
                />
              )}

              {/* 9. Effective Date — conditionally shown */}
              {fieldVisibility.effectiveDate && (
                <div className={styles.formGroup}>
                  <div>
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
                          setSubmitError("");
                          return;
                        }

                        if (isFutureDate(date)) {
                          setEffectiveDateError(FUTURE_EFFECTIVE_DATE_ERROR);
                          return;
                        }

                        setEffectiveDate(format(date, "yyyy-MM-dd"));
                        setEffectiveDateError("");
                        setSubmitError("");
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
                        return isValidDate(parsed) ? parsed : null;
                      }}
                      textField={{
                        errorMessage: effectiveDateError,
                        ariaLabel: "Effective Date",
                      }}
                      placeholder={
                        fieldRequirements.effectiveDate
                          ? "Select date (required)"
                          : "Select date (optional)"
                      }
                    />
                  </div>
                  <div style={{ fontSize: "12px", color: "#6b7280", marginTop: "4px" }}>
                    The date the document was aligned at TASC/AST.
                  </div>
                </div>
              )}

              {/* 10. Comments — conditionally shown */}
              {fieldVisibility.comments && (
                <div className={styles.formGroup}>
                  <TextField
                    label="Comments"
                    ariaLabel="Comments"
                    required={fieldRequirements.comments}
                    multiline
                    rows={2}
                    value={comments}
                    onChange={(_, val) =>
                      setComments((val || "").slice(0, COMMENTS_MAX_LENGTH))
                    }
                    placeholder={
                      fieldRequirements.comments
                        ? "Enter comments (required)"
                        : "Enter optional comments"
                    }
                    maxLength={COMMENTS_MAX_LENGTH}
                  />
                  <div className={styles.formCharCount}>
                    {comments.length}/{COMMENTS_MAX_LENGTH}
                  </div>
                </div>
              )}

              {/* 11. Description — conditionally shown */}
              {fieldVisibility.description && (
                <div className={styles.formGroup}>
                  <TextField
                    label="Description"
                    ariaLabel="Description"
                    required={fieldRequirements.description}
                    multiline
                    rows={3}
                    value={description}
                    onChange={(_, val) => setDescription(val || "")}
                    placeholder={
                      fieldRequirements.description
                        ? "Enter document description (required)"
                        : "Enter document description (optional)"
                    }
                  />
                </div>
              )}
            </>
          }

          <div className={styles.formNoticePlain}>
            Once a document is submitted, only an Admin can edit, archive or
            delete. Please check that the correct selections are made.
          </div>
        </div>

        <div className={styles.panelFooter}>
          {/* Upload Progress Indicator */}
          {uploadProgress && (
            <div style={{ width: "100%", marginBottom: "12px" }}>
              <div
                style={{
                  fontSize: "12px",
                  fontWeight: 500,
                  color: "#071d49",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  marginBottom: "8px",
                }}
              >
                {uploadProgress === "Completed" ? (
                  <>
                    <img
                      src={require("../../assets/icons/check.svg")}
                      alt=""
                      aria-hidden="true"
                      style={{ width: "14px", height: "14px", display: "inline-block" }}
                    />{" "}
                    Upload complete
                  </>
                ) : uploadProgress === "Failed" ? (
                  <>
                    <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true"
                      style={{ display: "inline-block", flexShrink: 0 }}>
                      <circle cx="7" cy="7" r="6" stroke="#dc2626" strokeWidth="1.5" />
                      <path d="M4.5 4.5L9.5 9.5M9.5 4.5L4.5 9.5" stroke="#dc2626"
                        strokeWidth="1.5" strokeLinecap="round" />
                    </svg>{" "}
                    Upload failed
                  </>
                ) : uploadProgress === "UploadingFile" ? (
                  <span>
                    Uploading file
                    {file && (
                      <span style={{ color: "var(--spectra-text-secondary)", fontWeight: 400 }}>
                        {" "}— {formatFileSize(Math.round(file.size * uploadPercent / 100))} of {formatFileSize(file.size)}
                      </span>
                    )}
                    <span style={{ marginLeft: 6 }}>{uploadPercent}%</span>
                  </span>
                ) : uploadProgress === "SavingMetadata" ? (
                  "Saving metadata..."
                ) : uploadProgress === "Preparing" ? (
                  "Preparing..."
                ) : (
                  "Finalizing..."
                )}
              </div>
              <div
                style={{
                  width: "100%",
                  height: "4px",
                  backgroundColor: "#e5e5e5",
                  borderRadius: "2px",
                  overflow: "hidden",
                }}
              >
                <div
                  style={{
                    height: "100%",
                    backgroundColor:
                      uploadProgress === "Failed"
                        ? "#dc2626"
                        : uploadProgress === "Completed"
                          ? "#16a34a"
                          : "#0f67e8",
                    width:
                      uploadProgress === "Preparing"
                        ? "5%"
                        : uploadProgress === "UploadingFile"
                          ? `${5 + Math.round(uploadPercent * 0.82)}%`
                          : uploadProgress === "SavingMetadata"
                            ? "90%"
                            : uploadProgress === "Finalizing"
                              ? "95%"
                              : "100%",
                    transition: uploadProgress === "UploadingFile" ? "width 0.1s linear" : "width 0.3s ease",
                  }}
                />
              </div>
            </div>
          )}
          <div
            style={{
              display: "flex",
              gap: "10px",
              justifyContent: "flex-end",
              width: "100%",
            }}
          >
            {isUploading ? (
              <button
                className={styles.btnSecondary}
                onClick={() => { cancelUpload?.(); }}
                aria-label="Cancel upload"
              >
                Cancel Upload
              </button>
            ) : (
              <button className={styles.btnSecondary} onClick={onCancel}>
                Cancel
              </button>
            )}
            <button
              className={styles.btnPrimary}
              onClick={handleSubmit}
              disabled={isUploading}
            >
              {isUploading ? (
                <>
                  <span className={styles.spinner} />{" "}
                  {uploadProgress === "SavingMetadata"
                    ? "Saving..."
                    : uploadProgress === "Finalizing"
                      ? "Finalizing..."
                      : "Uploading..."}
                </>
              ) : (
                <>
                  <img
                    src={require("../../assets/icons/submit.svg")}
                    alt=""
                    style={{ width: "16px", height: "16px", display: "inline-block", marginRight: "6px" }}
                    aria-hidden="true"
                  />
                  Submit
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </>
  );
};
