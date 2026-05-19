import * as React from "react";
import { IFilterState } from "../../interfaces/IFilterState";
import {
  IMetadataOptions,
  toValueArray,
} from "../../interfaces/IMetadataOptions";
import {
  getAllPaids,
  getDiseaseAreaStrategiesForTherapeuticArea,
} from "../../utils/cascadingFilterHelper";
import { DatePicker } from "@fluentui/react/lib/DatePicker";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { SearchableDropdown } from "../SearchableDropdown/SearchableDropdown";
import styles from "../SPECTRA.module.scss";

export interface IFilterPanelProps {
  isOpen: boolean;
  filters: IFilterState;
  resetToken: number;
  options: IMetadataOptions;
  onFilterChange: <K extends keyof IFilterState>(
    field: K,
    value: IFilterState[K],
  ) => void;
  onApply: () => void;
  onCancel: () => void;
  onReset: () => void;
}

export const FilterPanel: React.FC<IFilterPanelProps> = ({
  isOpen,
  filters,
  resetToken,
  options,
  onFilterChange,
  onApply,
  onCancel,
  onReset,
}) => {
  if (!isOpen) return null;

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

  const paidValues = getAllPaids(options.projectPaidRelationships);
  const indicationValues = React.useMemo(() => {
    const directValues = toValueArray(options.indications);
    if (directValues.length > 0) {
      return directValues;
    }

    return Array.from(
      new Set(
        options.projectPaidRelationships
          .map((relationship) =>
            typeof relationship.indication === "string"
              ? relationship.indication.trim()
              : "",
          )
          .filter((value) => value.length > 0),
      ),
    ).sort((a, b) => a.localeCompare(b));
  }, [options.indications, options.projectPaidRelationships]);

  const diseaseAreaStrategies = React.useMemo(
    () =>
      getDiseaseAreaStrategiesForTherapeuticArea(
        options.diseaseAreaStrategyRelationships,
        filters.therapeuticArea,
      ),
    [options.diseaseAreaStrategyRelationships, filters.therapeuticArea],
  );

  const effectiveDateValue = filters.effectiveDateFrom || undefined;

  React.useEffect(() => {
    if (filters.diseaseArea.length === 0) return;

    const validDiseaseAreas = filters.diseaseArea.filter((das) =>
      diseaseAreaStrategies.includes(das),
    );
    if (validDiseaseAreas.length !== filters.diseaseArea.length) {
      onFilterChange("diseaseArea", validDiseaseAreas);
    }
  }, [diseaseAreaStrategies, filters.diseaseArea, onFilterChange]);

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
          <span className={styles.panelTitle}>Filter</span>
          <div className={styles.panelHeaderActions}>
            <TooltipHost content="Reset all filters">
              <button
                className={styles.resetLink}
                onClick={onReset}
                aria-label="Reset all filters"
              >
                <img
                  src={require("../../assets/icons/reset.svg")}
                  alt=""
                  style={{
                    width: "16px",
                    height: "16px",
                    display: "inline-block",
                    marginRight: "4px",
                  }}
                  aria-hidden="true"
                />{" "}
                Reset
              </button>
            </TooltipHost>
          </div>
        </div>

        <div key={resetToken} className={styles.panelBody}>
          <div className={styles.formNoticePlain} style={{ marginTop: 0, marginBottom: 12, color: "#636363" }}>
            Filter dropdowns show current metadata values from source lists. Historical values on older documents may not appear here.
          </div>

          {/* Document Type */}
          <div className={styles.formGroup}>
            <SearchableDropdown
              label="Document Type"
              options={toValueArray(options.documentTypes)}
              selectedKeys={filters.documentType}
              onChange={(selected) => onFilterChange("documentType", selected)}
              placeholder="Type to search document type..."
              multiSelect={true}
            />
          </div>

          {/* Therapeutic Area */}
          <div className={styles.formGroup}>
            <SearchableDropdown
              label="Therapeutic Area"
              options={toValueArray(options.therapeuticAreas)}
              selectedKeys={filters.therapeuticArea}
              onChange={(selected) =>
                onFilterChange("therapeuticArea", selected)
              }
              placeholder="Type to search therapeutic area..."
              multiSelect={true}
            />
          </div>

          {/* DAS */}
          <SearchableDropdown
            label="DAS"
            options={diseaseAreaStrategies}
            selectedKeys={filters.diseaseArea}
            onChange={(selected) => onFilterChange("diseaseArea", selected)}
            placeholder="Type to search DAS..."
            multiSelect={true}
          />

          {/* Asset */}
          <SearchableDropdown
            label="Asset"
            options={toValueArray(options.assets)}
            selectedKeys={filters.asset}
            onChange={(selected) => onFilterChange("asset", selected)}
            placeholder="Type to search asset..."
            multiSelect={true}
          />

          {/* Indication */}
          <SearchableDropdown
            label="Indication"
            options={indicationValues}
            selectedKeys={filters.indication}
            onChange={(selected) => onFilterChange("indication", selected)}
            placeholder="Type to search indication..."
            multiSelect={true}
          />

          {/* PAID */}
          <SearchableDropdown
            label="PAID"
            options={paidValues}
            selectedKeys={filters.paid}
            onChange={(selected) => onFilterChange("paid", selected)}
            placeholder="Type to search PAID..."
            multiSelect={true}
          />

          {/* Effective Date */}
          <div>
            <div className={styles.formGroup}>
              <DatePicker
                label="Effective Date"
                value={effectiveDateValue}
                allowTextInput={true}
                calendarProps={{
                  calendarDayProps: {
                    styles: calendarDayStyles,
                  },
                }}
                onSelectDate={(date) => {
                  onFilterChange("effectiveDateFrom", date || null);
                  onFilterChange("effectiveDateTo", date || null);
                }}
                placeholder="Select effective date"
                ariaLabel="Effective date"
              />
              <div style={{ fontSize: "12px", color: "#6b7280", marginTop: "4px" }}>
                The date the document was aligned at TASC/AST.
              </div>
            </div>
          </div>
        </div>

        <div className={`${styles.panelFooter} ${styles.filterPanelFooter}`}>
          <button className={styles.btnSecondary} onClick={onCancel}>
            Cancel
          </button>
          <button className={styles.btnPrimary} onClick={onApply}>
            Apply
          </button>
        </div>
      </div>
    </>
  );
};
