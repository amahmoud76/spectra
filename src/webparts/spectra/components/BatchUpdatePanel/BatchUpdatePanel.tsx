import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  IMetadataOptions,
  toValueArray,
} from "../../interfaces/IMetadataOptions";
import { IFilterState } from "../../interfaces/IFilterState";
import {
  FIELD_NAMES,
  isFieldVisible,
  isFieldRequired,
  isFieldMultiSelect,
} from "../../config/fieldConfig";
import {
  getAllPaids,
  getDiseaseAreaStrategiesForTherapeuticArea,
} from "../../utils/cascadingFilterHelper";
import {
  BatchMetadataField,
  IBatchChangeSet,
} from "../../utils/batchUpdateHelper";
import { IUseBatchUpdateResult } from "../../hooks/useBatchUpdate";
import { SearchableDropdown } from "../SearchableDropdown/SearchableDropdown";
import { buildAssetSearchAliases } from "../../utils/assetSearchAliasHelper";
import { DateRangePicker } from "../DateRangePicker/DateRangePicker";
import { usePeopleSearch } from "../../hooks/usePeopleSearch";
import styles from "../SPECTRA.module.scss";

export interface IBatchUpdatePanelProps {
  isOpen: boolean;
  options: IMetadataOptions;
  context: WebPartContext;
  useMock: boolean;
  batch: IUseBatchUpdateResult;
  /** Collapse / overlay — closes the panel but keeps the working state. */
  onClose: () => void;
  /** Footer Cancel — resets the operation and closes. */
  onCancel: () => void;
  /** Navigate to the data table with the current criteria applied. */
  onPreviewList: () => void;
}

// Criteria (Step 1) metadata fields, in display order. Sub-TA omitted from Find.
const CRITERIA_FIELDS: Array<{
  field: keyof IFilterState;
  fieldName: string;
  label: string;
  placeholder: string;
  probeTA?: string;
}> = [
  {
    field: "therapeuticArea",
    fieldName: FIELD_NAMES.THERAPEUTIC_AREA,
    label: "Therapeutic Area",
    placeholder: "Type to search therapeutic area...",
  },
  {
    field: "diseaseArea",
    fieldName: FIELD_NAMES.DISEASE_AREA,
    label: "Disease Area",
    placeholder: "Type to search disease area...",
  },
  {
    field: "asset",
    fieldName: FIELD_NAMES.ASSET,
    label: "Asset",
    placeholder: "Type to search asset...",
  },
  {
    field: "indication",
    fieldName: FIELD_NAMES.INDICATION,
    label: "Indication",
    placeholder: "Type to search indication...",
  },
  {
    field: "lineOfTherapy",
    fieldName: FIELD_NAMES.LINE_OF_THERAPY,
    label: "Line of Therapy",
    placeholder: "Type to search line of therapy...",
    probeTA: "Oncology",
  },
  {
    field: "paid",
    fieldName: FIELD_NAMES.PAID,
    label: "PAID",
    placeholder: "Type to search PAID...",
  },
];

// Editable (Step 2) fields. Document Type is always disabled.
const EDITABLE_FIELDS: Array<{
  key: BatchMetadataField;
  fieldName: string;
  label: string;
  probeTA?: string;
}> = [
  {
    key: "therapeuticArea",
    fieldName: FIELD_NAMES.THERAPEUTIC_AREA,
    label: "Therapeutic Area",
  },
  {
    key: "subTherapeuticArea",
    fieldName: FIELD_NAMES.SUB_THERAPEUTIC_AREA,
    label: "Sub-Therapeutic Area",
    probeTA: "Aesthetics",
  },
  {
    key: "diseaseArea",
    fieldName: FIELD_NAMES.DISEASE_AREA,
    label: "Disease Area",
  },
  { key: "asset", fieldName: FIELD_NAMES.ASSET, label: "Asset" },
  { key: "indication", fieldName: FIELD_NAMES.INDICATION, label: "Indication" },
  {
    key: "lineOfTherapy",
    fieldName: FIELD_NAMES.LINE_OF_THERAPY,
    label: "Line of Therapy",
    probeTA: "Oncology",
  },
  { key: "paid", fieldName: FIELD_NAMES.PAID, label: "PAID" },
];

export const BatchUpdatePanel: React.FC<IBatchUpdatePanelProps> = ({
  isOpen,
  options,
  context,
  useMock,
  batch,
  onClose,
  onCancel,
  onPreviewList,
}) => {
  const {
    stage,
    criteria,
    matchedDocs,
    changeSet,
    isFinding,
    isProcessing,
    progress,
    result,
    setCriteria,
    findDocuments,
    backToCriteria,
    setChangeSet,
    computePlan,
    applyPlan,
    prepareNextRun,
    clearResult,
  } = batch;

  const {
    results: peopleResults,
    isSearching: isPeopleSearching,
    search: searchPeople,
    clear: clearPeopleResults,
  } = usePeopleSearch(context, useMock);
  const [peopleQuery, setPeopleQuery] = React.useState("");

  const documentType = criteria.documentType[0] || "";

  // ── Option value lists ─────────────────────────────────────
  const paidValues = getAllPaids(options.projectPaidRelationships);
  const indicationValues = React.useMemo(() => {
    const direct = toValueArray(options.indications);
    if (direct.length > 0) return direct;
    return Array.from(
      new Set(
        options.projectPaidRelationships
          .map((r) =>
            typeof r.indication === "string" ? r.indication.trim() : "",
          )
          .filter((v) => v.length > 0),
      ),
    ).sort((a, b) => a.localeCompare(b));
  }, [options.indications, options.projectPaidRelationships]);

  // Synonym search — typing any SEARCH_TOKEN keyword (e.g. "emraclidine")
  // matches the asset it belongs to (e.g. "ABBV-132").
  const assetSearchAliases = React.useMemo(
    () => buildAssetSearchAliases(options.searchTokenRows),
    [options.searchTokenRows],
  );

  const optionsForField = React.useCallback(
    (fieldName: string): string[] => {
      switch (fieldName) {
        case FIELD_NAMES.THERAPEUTIC_AREA:
          return toValueArray(options.therapeuticAreas);
        case FIELD_NAMES.SUB_THERAPEUTIC_AREA:
          return toValueArray(options.subTherapeuticAreas);
        case FIELD_NAMES.DISEASE_AREA:
          return getDiseaseAreaStrategiesForTherapeuticArea(
            options.diseaseAreaStrategyRelationships,
          );
        case FIELD_NAMES.ASSET:
          return toValueArray(options.assets);
        case FIELD_NAMES.INDICATION:
          return indicationValues;
        case FIELD_NAMES.LINE_OF_THERAPY:
          return toValueArray(options.lineOfTherapy);
        case FIELD_NAMES.PAID:
          return paidValues;
        default:
          return [];
      }
    },
    [options, indicationValues, paidValues],
  );

  // A field is applicable to the doc type if it could be visible under any TA.
  const isApplicable = React.useCallback(
    (fieldName: string, probeTA?: string): boolean => {
      if (!documentType) return false;
      return (
        isFieldVisible(fieldName, documentType) ||
        isFieldVisible(fieldName, documentType, probeTA || "Oncology") ||
        isFieldVisible(fieldName, documentType, "Aesthetics")
      );
    },
    [documentType],
  );

  // ── Criteria (Step 1) handlers ─────────────────────────────
  const updateCriteria = React.useCallback(
    <K extends keyof IFilterState>(field: K, value: IFilterState[K]) => {
      setCriteria({ ...criteria, [field]: value });
    },
    [criteria, setCriteria],
  );

  const handleDocTypeChange = React.useCallback(
    (docType: string) => {
      // Reset metadata criteria when the doc type changes; keep dates + createdBy.
      setCriteria({
        ...criteria,
        documentType: docType ? [docType] : [],
        therapeuticArea: [],
        subTherapeuticArea: [],
        diseaseArea: [],
        asset: [],
        indication: [],
        lineOfTherapy: [],
        paid: [],
      });
    },
    [criteria, setCriteria],
  );

  const handleFind = React.useCallback(() => {
    clearResult();
    void findDocuments();
  }, [clearResult, findDocuments]);

  // ── Values (Step 2) handlers ───────────────────────────────
  const isFieldChanged = (key: BatchMetadataField): boolean =>
    changeSet.metadata[key] !== undefined;
  const isEffectiveDateChanged = changeSet.effectiveDate !== undefined;

  const toggleField = React.useCallback(
    (key: BatchMetadataField, checked: boolean) => {
      const nextMeta = { ...changeSet.metadata };
      if (checked) nextMeta[key] = nextMeta[key] || [];
      else delete nextMeta[key];
      setChangeSet({ ...changeSet, metadata: nextMeta });
    },
    [changeSet, setChangeSet],
  );

  const setFieldValue = React.useCallback(
    (key: BatchMetadataField, values: string[]) => {
      setChangeSet({
        ...changeSet,
        metadata: { ...changeSet.metadata, [key]: values },
      });
    },
    [changeSet, setChangeSet],
  );

  const toggleEffectiveDate = React.useCallback(
    (checked: boolean) => {
      const next: IBatchChangeSet = { ...changeSet };
      if (checked) next.effectiveDate = next.effectiveDate ?? "";
      else delete next.effectiveDate;
      setChangeSet(next);
    },
    [changeSet, setChangeSet],
  );

  const setEffectiveDate = React.useCallback(
    (value: string) => {
      setChangeSet({ ...changeSet, effectiveDate: value });
    },
    [changeSet, setChangeSet],
  );

  // ── Validation ─────────────────────────────────────────────
  const validationErrors = React.useMemo((): string[] => {
    const errors: string[] = [];
    for (const f of EDITABLE_FIELDS) {
      if (!isFieldChanged(f.key)) continue;
      if (!isApplicable(f.fieldName, f.probeTA)) continue;
      const values = changeSet.metadata[f.key] || [];
      const required = isFieldRequired(f.fieldName, documentType, f.probeTA);
      const multi = isFieldMultiSelect(f.fieldName, documentType, f.probeTA);
      if (required && values.length === 0) {
        errors.push(`${f.label} is required and cannot be left blank.`);
      }
      if (!multi && values.length > 1) {
        errors.push(`${f.label} allows only a single value.`);
      }
    }
    if (isEffectiveDateChanged && !changeSet.effectiveDate) {
      errors.push("Effective Date is required and cannot be left blank.");
    }
    return errors;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [changeSet, documentType, isApplicable]);

  const hasSelectedChange =
    Object.keys(changeSet.metadata).length > 0 || isEffectiveDateChanged;

  const plan = React.useMemo(
    () => computePlan(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [computePlan],
  );

  const canApply =
    hasSelectedChange && validationErrors.length === 0 && !isProcessing;

  // ── Apply flow ─────────────────────────────────────────────
  // Applies directly — the plan's conflict/archive impact is already shown
  // inline in Step 2, so no extra confirmation modal is needed.
  const handleApply = React.useCallback(async () => {
    if (!canApply) return;
    const freshPlan = computePlan();
    await applyPlan(freshPlan);
    prepareNextRun();
  }, [canApply, computePlan, applyPlan, prepareNextRun]);

  // ── Criteria summary chips (Step 2) ────────────────────────
  const criteriaChips = React.useMemo((): string[] => {
    const chips: string[] = [];
    if (documentType) chips.push(`Document Type: ${documentType}`);
    const push = (label: string, values: string[]): void =>
      values.forEach((v) => chips.push(`${label}: ${v}`));
    push("Therapeutic Area", criteria.therapeuticArea);
    push("Disease Area", criteria.diseaseArea);
    push("Asset", criteria.asset);
    push("Indication", criteria.indication);
    push("Line of Therapy", criteria.lineOfTherapy);
    push("PAID", criteria.paid);
    push("Created By", criteria.createdBy);
    return chips;
  }, [criteria, documentType]);

  if (!isOpen) return null;

  return (
    <>
      <div className={styles.panelOverlay} onClick={onClose} />
      <div className={`${styles.panel} ${styles.panelWhiteBg}`}>
        <div className={styles.panelToggleBar}>
          <button
            className={styles.panelToggleBtn}
            onClick={onClose}
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
          <span className={styles.panelTitle}>Batch Update Metadata</span>
          <span
            style={{
              border: "1px solid #0F67E8",
              color: "#0F67E8",
              borderRadius: 0,
              padding: "2px 10px",
              fontSize: 12,
              fontWeight: 300,
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            <img
              src={require("../../assets/icons/shield.svg")}
              alt=""
              style={{ width: 14, height: 14 }}
              aria-hidden="true"
            />
            ADMIN TOOL
          </span>
        </div>

        <div className={styles.panelBody}>
          {/* Success banner (after a successful apply) */}
          {result && (
            <div
              className={styles.panelSuccessNotice}
              role="status"
              style={{ marginBottom: 16 }}
            >
              <img
                src={require("../../assets/icons/check.svg")}
                alt=""
                aria-hidden="true"
                style={{
                  width: 18,
                  height: 18,
                  display: "block",
                  flexShrink: 0,
                }}
              />
              <span>
                {result.updated} document(s) have been updated
                {result.archived > 0
                  ? `, ${result.archived} archived as duplicate(s)`
                  : ""}
                {result.failed > 0 ? `, ${result.failed} skipped` : ""}.
              </span>
            </div>
          )}

          {stage === "criteria" ? renderCriteria() : renderValues()}
        </div>

        <div
          className={styles.panelFooter}
          style={{ justifyContent: "space-between" }}
        >
          <button className={styles.btnSecondary} onClick={onCancel}>
            Cancel
          </button>
          {stage === "criteria" ? (
            <button
              className={styles.btnPrimary}
              onClick={handleFind}
              disabled={!documentType || isFinding}
            >
              <img
                src={require("../../assets/icons/magnifying-glass.svg")}
                alt=""
                aria-hidden="true"
                style={{
                  width: 14,
                  height: 14,
                  filter: "brightness(0) invert(1)",
                }}
              />
              {isFinding ? "Finding…" : "Find Documents"}
            </button>
          ) : (
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                alignItems: "flex-end",
                gap: 6,
                flex: 1,
              }}
            >
              {isProcessing ? (
                <div className={styles.batchProgress} aria-live="polite">
                  <div className={styles.batchProgressHeader}>
                    <span className={styles.spinner} />
                    <span>
                      Updating {progress ? progress.done : 0} of{" "}
                      {progress ? progress.total : 0}…
                    </span>
                  </div>
                  <div className={styles.batchProgressTrack}>
                    <div
                      className={styles.batchProgressFill}
                      style={{
                        width:
                          progress && progress.total > 0
                            ? `${Math.round((progress.done / progress.total) * 100)}%`
                            : "0%",
                      }}
                    />
                  </div>
                  {progress?.currentFileName && (
                    <div
                      className={styles.batchProgressFile}
                      title={progress.currentFileName}
                    >
                      {progress.currentFileName}
                    </div>
                  )}
                </div>
              ) : (
                <>
                  <button
                    className={styles.btnPrimary}
                    onClick={() => void handleApply()}
                    disabled={!canApply}
                  >
                    ⚠ Apply Update
                  </button>
                  <div style={{ fontSize: 12, color: "#6b7280" }}>
                    This action cannot be undone. Verify selections before
                    applying.
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </>
  );

  // ─────────────────────────────────────────────────────────
  // Render: Step 1 — Find documents
  // ─────────────────────────────────────────────────────────
  function renderCriteria(): JSX.Element {
    return (
      <>
        <div
          className={styles.formNoticePlain}
          style={{ marginTop: 0, marginBottom: 12, color: "#636363" }}
        >
          <strong>STEP 1: FIND DOCUMENTS BASED ON METADATA TAG</strong>
          <br />
          The system searches documents that match all of the fields you select.
          If you pick more than one option within a field, it matches documents
          with one or the other value.
        </div>

        {/* Document Type — single-select, drives the contract */}
        <div className={styles.formGroup}>
          <label className={styles.formLabel}>Document Type</label>
          <select
            className={styles.formInput}
            value={documentType}
            onChange={(e) => handleDocTypeChange(e.target.value)}
            aria-label="Document Type"
          >
            <option value="">Select</option>
            {toValueArray(options.documentTypes).map((dt) => (
              <option key={dt} value={dt}>
                {dt}
              </option>
            ))}
          </select>
        </div>

        {documentType &&
          CRITERIA_FIELDS.filter((f) =>
            isApplicable(f.fieldName, f.probeTA),
          ).map((f) => (
            <div className={styles.formGroup} key={f.field}>
              <SearchableDropdown
                label={f.label}
                options={optionsForField(f.fieldName)}
                selectedKeys={criteria[f.field] as string[]}
                onChange={(selected) => updateCriteria(f.field, selected)}
                placeholder={f.placeholder}
                multiSelect={true}
                showChipsBelow={true}
                searchAliases={
                  f.fieldName === FIELD_NAMES.ASSET
                    ? assetSearchAliases
                    : undefined
                }
              />
            </div>
          ))}

        {documentType && (
          <>
            <div className={styles.dateRangeRow}>
              <DateRangePicker
                label="Effective Date Range"
                fromDate={criteria.effectiveDateFrom}
                toDate={criteria.effectiveDateTo}
                onFromChange={(d) => updateCriteria("effectiveDateFrom", d)}
                onToChange={(d) => updateCriteria("effectiveDateTo", d)}
              />
              <DateRangePicker
                label="Uploaded Files Date Range"
                fromDate={criteria.uploadDateFrom}
                toDate={criteria.uploadDateTo}
                onFromChange={(d) => updateCriteria("uploadDateFrom", d)}
                onToChange={(d) => updateCriteria("uploadDateTo", d)}
              />
            </div>

            {/* Created By — async AD people search */}
            <div className={styles.formGroup}>
              <label className={styles.formLabel}>Created By</label>
              <div className={styles.peoplePickerWrap}>
                <input
                  type="text"
                  className={styles.formInput}
                  placeholder="Type a name to search…"
                  value={peopleQuery}
                  onChange={(e) => {
                    setPeopleQuery(e.target.value);
                    searchPeople(e.target.value);
                  }}
                  onBlur={() => setTimeout(() => clearPeopleResults(), 200)}
                  aria-label="Search for document creator"
                />
                {(isPeopleSearching || peopleResults.length > 0) &&
                  peopleQuery.trim().length >= 2 && (
                    <div className={styles.peoplePickerSuggestions}>
                      {isPeopleSearching ? (
                        <div className={styles.peoplePickerStatus}>
                          Searching…
                        </div>
                      ) : (
                        peopleResults
                          .filter(
                            (r) => !criteria.createdBy.includes(r.displayName),
                          )
                          .map((person) => (
                            <div
                              key={person.email || person.displayName}
                              className={styles.peoplePickerItem}
                              onMouseDown={(e) => e.preventDefault()}
                              onClick={() => {
                                updateCriteria("createdBy", [
                                  ...criteria.createdBy,
                                  person.displayName,
                                ]);
                                setPeopleQuery("");
                                clearPeopleResults();
                              }}
                            >
                              <span className={styles.peoplePickerItemName}>
                                {person.displayName}
                              </span>
                              {person.email && (
                                <span className={styles.peoplePickerItemEmail}>
                                  {person.email}
                                </span>
                              )}
                            </div>
                          ))
                      )}
                    </div>
                  )}
              </div>
              {criteria.createdBy.length > 0 && (
                <div className={styles.nativeMultiSelectChips}>
                  {criteria.createdBy.map((name) => (
                    <span key={name} className={styles.chip}>
                      {name}
                      <button
                        className={styles.chipRemove}
                        onClick={() =>
                          updateCriteria(
                            "createdBy",
                            criteria.createdBy.filter((n) => n !== name),
                          )
                        }
                        aria-label={`Remove ${name}`}
                        type="button"
                      >
                        ×
                      </button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          </>
        )}
      </>
    );
  }

  // ─────────────────────────────────────────────────────────
  // Render: Step 2 — Apply metadata change
  // ─────────────────────────────────────────────────────────
  function renderValues(): JSX.Element {
    return (
      <>
        <div
          style={{
            border: "1px solid #e5e7eb",
            borderRadius: 8,
            padding: 16,
            marginBottom: 16,
          }}
        >
          <button
            className={styles.btnSecondary}
            onClick={backToCriteria}
            style={{
              marginBottom: 12,
              minHeight: 28,
              padding: "0 14px",
              fontSize: 13,
            }}
          >
            ← Back
          </button>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 12,
              flexWrap: "wrap",
            }}
          >
            <strong>Total Documents Found:</strong>
            <span
              style={{
                background: "#5B3FD4",
                color: "#fff",
                borderRadius: 6,
                padding: "2px 10px",
                fontWeight: 600,
              }}
            >
              {matchedDocs.length} items
            </span>
            <button
              type="button"
              onClick={onPreviewList}
              style={{
                background: "none",
                border: "none",
                color: "var(--spectra-action-blue)",
                cursor: "pointer",
                fontWeight: 600,
              }}
            >
              Preview List →
            </button>
          </div>
          {criteriaChips.length > 0 && (
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: 8,
                marginTop: 12,
              }}
            >
              {criteriaChips.map((c) => (
                <span
                  key={c}
                  className={styles.chip}
                  style={{ border: "1px solid #5B3FD4", color: "#5B3FD4" }}
                >
                  {c}
                </span>
              ))}
            </div>
          )}
        </div>

        <div className={styles.formNoticePlain} style={{ marginBottom: 12 }}>
          <strong>STEP 2: APPLY METADATA CHANGE</strong>
        </div>

        {/* Document Type — always disabled */}
        {renderFieldRow(
          "documentType-row",
          "Document Type",
          false,
          true,
          () => undefined,
          null,
        )}

        {EDITABLE_FIELDS.map((f) => {
          const applicable = isApplicable(f.fieldName, f.probeTA);
          const checked = isFieldChanged(f.key);
          const multi = isFieldMultiSelect(
            f.fieldName,
            documentType,
            f.probeTA,
          );
          return renderFieldRow(
            f.key,
            f.label,
            checked,
            !applicable,
            (next) => toggleField(f.key, next),
            checked && applicable ? (
              <div style={{ marginTop: 10 }}>
                <div
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    color: "#374151",
                    marginBottom: 4,
                  }}
                >
                  TO (New Target Value)
                </div>
                <SearchableDropdown
                  label=""
                  options={optionsForField(f.fieldName)}
                  selectedKeys={changeSet.metadata[f.key] || []}
                  onChange={(selected) => setFieldValue(f.key, selected)}
                  placeholder="Select value(s)…"
                  multiSelect={multi}
                  showChipsBelow={true}
                  searchAliases={
                    f.fieldName === FIELD_NAMES.ASSET
                      ? assetSearchAliases
                      : undefined
                  }
                />
              </div>
            ) : null,
          );
        })}

        {/* Effective Date */}
        {renderFieldRow(
          "effectiveDate-row",
          "Effective Date",
          isEffectiveDateChanged,
          false,
          (next) => toggleEffectiveDate(next),
          isEffectiveDateChanged ? (
            <div style={{ marginTop: 10 }}>
              <div
                style={{
                  fontSize: 12,
                  fontWeight: 600,
                  color: "#374151",
                  marginBottom: 4,
                }}
              >
                TO (New Target Value)
              </div>
              <input
                type="date"
                className={styles.formInput}
                value={changeSet.effectiveDate || ""}
                max={new Date().toISOString().split("T")[0]}
                onChange={(e) => {
                  const val = e.target.value;
                  if (
                    val &&
                    new Date(val) > new Date(new Date().toDateString())
                  ) {
                    return; // silently block future dates (browser enforces max attr)
                  }
                  setEffectiveDate(val);
                }}
                aria-label="New effective date"
              />
            </div>
          ) : null,
        )}

        {/* Bottom status banner */}
        {hasSelectedChange && (
          <div
            style={{
              marginTop: 16,
              border: `1px solid ${plan.hasConflicts ? "#DC2626" : "#5B3FD4"}`,
              color: plan.hasConflicts ? "#DC2626" : "#5B3FD4",
              borderRadius: 8,
              padding: "12px 16px",
              display: "flex",
              alignItems: "center",
              gap: 10,
              fontWeight: 600,
              fontSize: 16,
              fontFamily: "var(--spectra-font-heading)",
            }}
            role={plan.hasConflicts ? "alert" : "status"}
          >
            {plan.hasConflicts ? (
              <span aria-hidden="true">⚠</span>
            ) : (
              <img
                src={require("../../assets/icons/Information.svg")}
                alt=""
                aria-hidden="true"
                style={{ width: 16, height: 16, flexShrink: 0 }}
              />
            )}
            {plan.hasConflicts
              ? `Of the selected ${matchedDocs.length} documents, ${plan.toArchiveCount} will be archived based on duplicate metadata.`
              : `${plan.toUpdateCount} document(s) will be updated.`}
          </div>
        )}

        {/* Progress while applying */}
        {isProcessing && progress && (
          <div style={{ marginTop: 12, fontSize: 13, color: "#374151" }}>
            Updating {progress.done} of {progress.total}…{" "}
            {progress.currentFileName}
          </div>
        )}
      </>
    );
  }

  function renderFieldRow(
    key: string,
    label: string,
    checked: boolean,
    disabled: boolean,
    onToggle: (checked: boolean) => void,
    body: React.ReactNode,
  ): JSX.Element {
    return (
      <div
        key={key}
        style={{
          border: checked ? "1px solid #5B3FD4" : "1px solid #e5e7eb",
          borderRadius: 8,
          padding: 16,
          marginBottom: 12,
          opacity: disabled ? 0.5 : 1,
        }}
      >
        <label
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            cursor: disabled ? "not-allowed" : "pointer",
            fontWeight: 600,
            color: checked ? "#5B3FD4" : "#111827",
          }}
        >
          <input
            type="checkbox"
            checked={checked}
            disabled={disabled}
            onChange={(e) => onToggle(e.target.checked)}
          />
          {label}
        </label>
        {body}
      </div>
    );
  }
};
