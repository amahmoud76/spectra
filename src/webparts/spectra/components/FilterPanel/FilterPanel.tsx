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
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
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

  const toDropdownOptions = (values: string[]): IDropdownOption[] =>
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

  const selectedTherapeuticArea = filters.therapeuticArea[0];
  const diseaseAreaStrategies = React.useMemo(
    () =>
      getDiseaseAreaStrategiesForTherapeuticArea(
        options.diseaseAreaStrategyRelationships,
        selectedTherapeuticArea,
      ),
    [options.diseaseAreaStrategyRelationships, selectedTherapeuticArea],
  );

  const effectiveDateValue = filters.effectiveDateFrom || undefined;

  React.useEffect(() => {
    if (filters.diseaseArea.length === 0) return;

    const selectedDas = filters.diseaseArea[0];
    if (!diseaseAreaStrategies.includes(selectedDas)) {
      onFilterChange("diseaseArea", []);
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
            <Dropdown
              label="Document Type"
              options={toDropdownOptions(
                toValueArray(options.documentTypes),
              )}
              selectedKey={
                filters.documentType.length > 0
                  ? filters.documentType[0]
                  : undefined
              }
              onChange={(_, item) =>
                onFilterChange("documentType", item ? [item.key as string] : [])
              }
              placeholder="Select document type"
            />
          </div>

          {/* Therapeutic Area */}
          <div className={styles.formGroup}>
            <Dropdown
              label="Therapeutic Area"
              options={toDropdownOptions(
                toValueArray(options.therapeuticAreas),
              )}
              selectedKey={
                filters.therapeuticArea.length > 0
                  ? filters.therapeuticArea[0]
                  : undefined
              }
              onChange={(_, item) =>
                onFilterChange(
                  "therapeuticArea",
                  item ? [item.key as string] : [],
                )
              }
              placeholder="Select therapeutic area"
            />
          </div>

          {/* DAS */}
          <SearchableDropdown
            label="DAS"
            options={diseaseAreaStrategies}
            selectedKeys={filters.diseaseArea}
            onChange={(selected) => onFilterChange("diseaseArea", selected)}
            placeholder="Type to search DAS..."
            multiSelect={false}
          />

          {/* Asset */}
          <SearchableDropdown
            label="Asset"
            options={toValueArray(options.assets)}
            selectedKeys={filters.asset}
            onChange={(selected) => onFilterChange("asset", selected)}
            placeholder="Type to search asset..."
            multiSelect={false}
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
            multiSelect={false}
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
