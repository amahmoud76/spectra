import * as React from "react";
import { format } from "date-fns";
import { IFilterState } from "../../interfaces/IFilterState";
import styles from "../SPECTRA.module.scss";

export interface IActiveFilterChipsProps {
  filters: IFilterState;
  onRemoveValue: (field: keyof IFilterState, value: string) => void;
  onRemoveDateRange: (type: "effective" | "upload") => void;
  onClearAll: () => void;
}

type ArrayFilterField =
  | "documentType"
  | "therapeuticArea"
  | "subTherapeuticArea"
  | "asset"
  | "indication"
  | "lineOfTherapy"
  | "paid"
  | "diseaseArea";

const FIELD_LABELS: Record<ArrayFilterField, string> = {
  documentType: "Doc Type",
  therapeuticArea: "TA",
  subTherapeuticArea: "Sub TA",
  asset: "Asset",
  indication: "Indication",
  lineOfTherapy: "Line of Therapy",
  paid: "PAID",
  diseaseArea: "DAS",
};

const ARRAY_FIELDS: ArrayFilterField[] = [
  "documentType",
  "therapeuticArea",
  "subTherapeuticArea",
  "asset",
  "indication",
  "lineOfTherapy",
  "paid",
  "diseaseArea",
];

const formatDate = (date: Date): string => {
  try {
    return format(date, "MMM d, yyyy");
  } catch {
    return "";
  }
};

export const ActiveFilterChips: React.FC<IActiveFilterChipsProps> = ({
  filters,
  onRemoveValue,
  onRemoveDateRange,
  onClearAll,
}) => {
  const chips: React.ReactNode[] = [];

  for (const field of ARRAY_FIELDS) {
    const values = filters[field] as string[];
    if (values.length === 0) continue;
    const prefix = FIELD_LABELS[field];
    values.forEach((value) => {
      chips.push(
        <span key={`${field}-${value}`} className={styles.activeFilterChip}>
          <span className={styles.activeFilterChipLabel}>
            <span className={styles.activeFilterChipPrefix}>{prefix}:</span>{" "}
            {value}
          </span>
          <button
            className={styles.activeFilterChipRemove}
            onClick={() => onRemoveValue(field, value)}
            aria-label={`Remove ${prefix}: ${value}`}
            type="button"
          >
            ×
          </button>
        </span>,
      );
    });
  }

  if (filters.effectiveDateFrom || filters.effectiveDateTo) {
    const from = filters.effectiveDateFrom
      ? formatDate(filters.effectiveDateFrom)
      : "";
    const to = filters.effectiveDateTo
      ? formatDate(filters.effectiveDateTo)
      : "";
    const label =
      from && to && from === to ? from : [from, to].filter(Boolean).join(" – ");
    chips.push(
      <span key="effectiveDate" className={styles.activeFilterChip}>
        <span className={styles.activeFilterChipLabel}>
          <span className={styles.activeFilterChipPrefix}>Effective Date:</span>{" "}
          {label}
        </span>
        <button
          className={styles.activeFilterChipRemove}
          onClick={() => onRemoveDateRange("effective")}
          aria-label="Remove effective date filter"
          type="button"
        >
          ×
        </button>
      </span>,
    );
  }

  if (filters.uploadDateFrom || filters.uploadDateTo) {
    const from = filters.uploadDateFrom
      ? formatDate(filters.uploadDateFrom)
      : "";
    const to = filters.uploadDateTo ? formatDate(filters.uploadDateTo) : "";
    const label =
      from && to && from === to ? from : [from, to].filter(Boolean).join(" – ");
    chips.push(
      <span key="uploadDate" className={styles.activeFilterChip}>
        <span className={styles.activeFilterChipLabel}>
          <span className={styles.activeFilterChipPrefix}>Upload Date:</span>{" "}
          {label}
        </span>
        <button
          className={styles.activeFilterChipRemove}
          onClick={() => onRemoveDateRange("upload")}
          aria-label="Remove upload date filter"
          type="button"
        >
          ×
        </button>
      </span>,
    );
  }

  if (chips.length === 0) return null;

  return (
    <div
      className={styles.activeFilterChipsBar}
      role="group"
      aria-label="Active filters"
    >
      <span className={styles.activeFilterChipsBarLabel}>Filters:</span>
      {chips}
      {chips.length > 1 && (
        <button
          className={styles.activeFilterClearAll}
          onClick={onClearAll}
          type="button"
        >
          Clear all
        </button>
      )}
    </div>
  );
};
