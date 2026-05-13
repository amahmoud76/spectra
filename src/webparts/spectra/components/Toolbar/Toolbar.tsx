import * as React from "react";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import styles from "../SPECTRA.module.scss";

export interface IToolbarProps {
  role: EffectiveRole;
  activeFilterCount: number;
  onUploadClick?: () => void;
  onFilterClick?: () => void;
  onClearFiltersClick?: () => void;
  onExportClick?: () => void;
  showUpload?: boolean;
  showFilter?: boolean;
  showClearFilters?: boolean;
  showExport?: boolean;
}

export const Toolbar: React.FC<IToolbarProps> = ({
  role,
  activeFilterCount,
  onUploadClick,
  onFilterClick,
  onClearFiltersClick,
  onExportClick,
  showUpload = true,
  showFilter = true,
  showClearFilters = true,
  showExport = false,
}) => {
  return (
    <div className={styles.toolbar}>
      {/* Upload — Admin and Contributor only */}
      {showUpload && role !== "viewer" && (
        <TooltipHost content="Upload a new document">
          <button
            className={styles.btnPrimary}
            onClick={onUploadClick}
            aria-label="Upload document"
            type="button"
          >
            <img
              src={require("../../assets/icons/upload.svg")}
              alt=""
              style={{
                width: "16px",
                height: "16px",
                display: "inline-block",
                marginRight: "6px",
              }}
              aria-hidden="true"
            />
            Upload Document
          </button>
        </TooltipHost>
      )}

      {/* Filter — All roles */}
      {showFilter && (
        <TooltipHost content="Filter documents">
          <button
            className={styles.btnSecondary}
            onClick={onFilterClick}
            aria-label={`Filter documents${activeFilterCount > 0 ? ` (${activeFilterCount} active)` : ""}`}
            type="button"
          >
            <img
              src={require("../../assets/icons/filter.svg")}
              alt=""
              style={{ width: "16px", height: "16px", display: "inline-block" }}
              aria-hidden="true"
            />
            Filter
            {activeFilterCount > 0 && (
              <span className={styles.filterBadge}>{activeFilterCount}</span>
            )}
          </button>
        </TooltipHost>
      )}

      {/* Clear Filters — All roles when filters are active */}
      {showClearFilters && activeFilterCount > 0 && (
        <TooltipHost content="Clear all active filters">
          <button
            className={styles.btnSecondary}
            onClick={onClearFiltersClick}
            aria-label={`Clear ${activeFilterCount} active filter${activeFilterCount > 1 ? "s" : ""}`}
            type="button"
          >
            Clear Filters
          </button>
        </TooltipHost>
      )}

      {/* Export CSV — Admin only */}
      {showExport && role === "admin" && (
        <TooltipHost content="Export documents to CSV">
          <button
            className={styles.btnSecondary}
            onClick={onExportClick}
            aria-label="Export documents to CSV"
            type="button"
          >
            <i className="fa fa-file-excel" />
            Export CSV
          </button>
        </TooltipHost>
      )}
    </div>
  );
};
