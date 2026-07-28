import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFilterState } from "../../interfaces/IFilterState";
import {
  IMetadataOptions,
  toValueArray,
} from "../../interfaces/IMetadataOptions";
import {
  getAllPaids,
  getDiseaseAreaStrategiesForTherapeuticArea,
} from "../../utils/cascadingFilterHelper";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { SearchableDropdown } from "../SearchableDropdown/SearchableDropdown";
import { buildAssetSearchAliases } from "../../utils/assetSearchAliasHelper";
import { DateRangePicker } from "../DateRangePicker/DateRangePicker";
import { usePeopleSearch } from "../../hooks/usePeopleSearch";
import styles from "../SPECTRA.module.scss";

export interface IFilterPanelProps {
  isOpen: boolean;
  filters: IFilterState;
  resetToken: number;
  options: IMetadataOptions;
  context: WebPartContext;
  useMock: boolean;
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
  context,
  useMock,
  onFilterChange,
  onApply,
  onCancel,
  onReset,
}) => {
  const {
    results: peopleResults,
    isSearching: isPeopleSearching,
    search: searchPeople,
    clear: clearPeopleResults,
  } = usePeopleSearch(context, useMock);
  const [peopleQuery, setPeopleQuery] = React.useState("");

  // Reset people query when panel closes or filters are reset
  React.useEffect(() => {
    if (!isOpen) {
      setPeopleQuery("");
      clearPeopleResults();
    }
  }, [isOpen, clearPeopleResults]);

  React.useEffect(() => {
    setPeopleQuery("");
    clearPeopleResults();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [resetToken]);

  // Synonym search — typing any SEARCH_TOKEN keyword (e.g. "emraclidine")
  // matches the asset it belongs to (e.g. "ABBV-132").
  const assetSearchAliases = React.useMemo(
    () => buildAssetSearchAliases(options.searchTokenRows),
    [options.searchTokenRows],
  );

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

  // Tracks the DAS cascade options at the last moment the DAS field was empty.
  // Auto-select only fires when those options changed (TA changed), not on manual clears.
  const dasCascadeRef = React.useRef<string[]>([]);

  React.useEffect(() => {
    if (filters.diseaseArea.length === 0) return;
    const validDiseaseAreas = filters.diseaseArea.filter((das) =>
      diseaseAreaStrategies.includes(das),
    );
    if (validDiseaseAreas.length !== filters.diseaseArea.length) {
      onFilterChange("diseaseArea", validDiseaseAreas);
    }
  }, [diseaseAreaStrategies, filters.diseaseArea, onFilterChange]);

  React.useEffect(() => {
    const cascade = diseaseAreaStrategies;
    if (filters.diseaseArea.length === 0) {
      if (
        filters.therapeuticArea.length > 0 &&
        cascade.join(",") !== dasCascadeRef.current.join(",") &&
        cascade.length === 1
      ) {
        onFilterChange("diseaseArea", [cascade[0]]);
      }
      dasCascadeRef.current = cascade;
    }
  }, [
    diseaseAreaStrategies,
    filters.diseaseArea,
    filters.therapeuticArea,
    onFilterChange,
  ]);

  // All hooks above — safe to do early return now
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
          <div
            className={styles.formNoticePlain}
            style={{ marginTop: 0, marginBottom: 12, color: "#636363" }}
          >
            Filter dropdowns show current metadata values from source lists.
            Historical values on older documents may not appear here.
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
            searchAliases={assetSearchAliases}
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

          {/* Date Range Filters */}
          <div className={styles.dateRangeRow}>
            <DateRangePicker
              label="Effective Date Range"
              fromDate={filters.effectiveDateFrom}
              toDate={filters.effectiveDateTo}
              onFromChange={(d) => onFilterChange("effectiveDateFrom", d)}
              onToChange={(d) => onFilterChange("effectiveDateTo", d)}
            />
            <DateRangePicker
              label="Uploaded Files Date Range"
              fromDate={filters.uploadDateFrom}
              toDate={filters.uploadDateTo}
              onFromChange={(d) => onFilterChange("uploadDateFrom", d)}
              onToChange={(d) => onFilterChange("uploadDateTo", d)}
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
                          (r) => !filters.createdBy.includes(r.displayName),
                        )
                        .map((person) => (
                          <div
                            key={person.email || person.displayName}
                            className={styles.peoplePickerItem}
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => {
                              onFilterChange("createdBy", [
                                ...filters.createdBy,
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

            {filters.createdBy.length > 0 && (
              <div className={styles.nativeMultiSelectChips}>
                {filters.createdBy.map((name) => (
                  <span key={name} className={styles.chip}>
                    {name}
                    <button
                      className={styles.chipRemove}
                      onClick={() =>
                        onFilterChange(
                          "createdBy",
                          filters.createdBy.filter((n) => n !== name),
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
