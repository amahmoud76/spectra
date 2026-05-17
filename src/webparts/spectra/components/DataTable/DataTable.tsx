import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { ISortState, SortField } from "../../interfaces/ISortState";
import { SearchMatchKind } from "../../utils/filterHelper";
import {
  FILE_NAME_DISPLAY_MAX_LENGTH,
  FILE_NAME_DISPLAY_MAX_LENGTH_WITH_BADGE,
  truncateFileNameForDisplay,
} from "../../utils/fileHelper";
import { SearchMatchBadge } from "../SearchMatchBadge/SearchMatchBadge";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { parseISO, format, isValid } from "date-fns";
import styles from "../SPECTRA.module.scss";

export interface IDataTableProps {
  documents: IDocument[];
  role: EffectiveRole;
  sortState: ISortState | null;
  onSort: (field: SortField) => void;
  onDocumentClick: (doc: IDocument) => void;
  onEditClick?: (doc: IDocument) => void;
  onArchiveClick?: (doc: IDocument) => void;
  onDeleteClick?: (doc: IDocument) => void;
  onArchiveReplaceClick?: (
    doc: IDocument,
    anchorPosition?: { top: number; left: number },
  ) => void;
  onReActivateClick?: (doc: IDocument) => void;
  isLoading: boolean;
  searchMatchKindByDocumentId?: Map<string, SearchMatchKind>;
  useEnhancedStyle?: boolean;
}

const TYPE_CHIP_COLORS: Record<string, { bg: string; color: string }> = {
  DAS:  { bg: "#7C3AED", color: "#fff" },
  TPP:  { bg: "#0891B2", color: "#fff" },
  TPC:  { bg: "#2563EB", color: "#fff" },
  IEP:  { bg: "#16A34A", color: "#fff" },
  EBP:  { bg: "#EA580C", color: "#fff" },
  EIVP: { bg: "#BE185D", color: "#fff" },
  IAS:  { bg: "#374151", color: "#fff" },
};
const TYPE_CHIP_DEFAULT = { bg: "#6B7280", color: "#fff" };

interface IColumnDef {
  key: string;
  label: string;
  sortField?: SortField;
  getValue: (doc: IDocument) => string;
  truncate?: boolean;
  adminOnly?: boolean;
}

const formatDate = (iso: string): string => {
  if (!iso) return "";
  const date = parseISO(iso);
  if (!isValid(date)) return iso;
  return format(date, "MMM/dd/yyyy");
};

const COLUMN_CLASS_BY_KEY: Record<string, string> = {
  fileName: "colFileName",
  asset: "colAsset",
  type: "colType",
  ta: "colTa",
  indication: "colIndication",
  paid: "colPaid",
  diseaseArea: "colDiseaseArea",
  date: "colDate",
  uploadDate: "colUploadDate",
  createdBy: "colCreatedBy",
  modifiedBy: "colModifiedBy",
  comments: "colComments",
  status: "colStatus",
};

const COLUMNS: IColumnDef[] = [
  {
    key: "fileName",
    label: "Document Name",
    sortField: "fileName",
    getValue: (d) => d.fileName,
    truncate: true,
  },
  { key: "asset", label: "Asset", getValue: (d) => d.asset.join("; ") },
  {
    key: "type",
    label: "Type",
    sortField: "documentType",
    getValue: (d) => d.documentType,
  },
  { key: "ta", label: "TA", getValue: (d) => d.therapeuticArea.join("; ") },
  {
    key: "indication",
    label: "Indication",
    getValue: (d) => d.indication.join("; "),
  },
  { key: "paid", label: "PAID", getValue: (d) => d.paid.join("; ") },
  {
    key: "diseaseArea",
    label: "Disease Area",
    sortField: "diseaseArea",
    getValue: (d) => d.diseaseArea.join("; "),
  },
  {
    key: "date",
    label: "Eff Date",
    sortField: "effectiveDate",
    getValue: (d) => formatDate(d.effectiveDate),
  },
  {
    key: "uploadDate",
    label: "Upload Date",
    sortField: "uploadDate",
    getValue: (d) => formatDate(d.uploadDate),
    adminOnly: true,
  },
  {
    key: "createdBy",
    label: "Created By",
    sortField: "createdBy",
    getValue: (d) => d.createdBy,
    truncate: true,
    adminOnly: true,
  },
  {
    key: "modifiedBy",
    label: "Last Modified By",
    sortField: "modifiedBy",
    getValue: (d) => d.modifiedBy,
    truncate: true,
    adminOnly: true,
  },
  {
    key: "comments",
    label: "Comments",
    sortField: "comments",
    getValue: (d) => d.comments,
  },
  {
    key: "status",
    label: "Status",
    sortField: "status",
    getValue: (d) => d.status,
    adminOnly: true,
  },
];

export const DataTable: React.FC<IDataTableProps> = ({
  documents,
  role,
  sortState,
  onSort,
  onDocumentClick,
  onEditClick,
  onArchiveClick,
  onDeleteClick,
  onArchiveReplaceClick,
  onReActivateClick,
  isLoading,
  searchMatchKindByDocumentId,
  useEnhancedStyle = false,
}) => {
  const getAnchorPosition = React.useCallback(
    (event: React.MouseEvent<HTMLElement>): { top: number; left: number } => {
      const rect = event.currentTarget.getBoundingClientRect();
      return {
        top: rect.bottom + 8,
        left: rect.left,
      };
    },
    [],
  );

  const visibleColumns = COLUMNS.filter((col) => {
    if (col.adminOnly && role !== "admin") return false;
    return true;
  });

  const styleClassMap = styles as unknown as Record<string, string>;

  const getColumnClassName = (key: string): string => {
    const classKey = COLUMN_CLASS_BY_KEY[key];
    return classKey ? styleClassMap[classKey] || "" : "";
  };

  const renderSortIcon = (field?: SortField): React.ReactNode => {
    if (!field) return null;

    const isActive = sortState && sortState.field === field;
    const isAsc = isActive && sortState.direction === "asc";

    return (
      <img
        src={require('../../assets/icons/sorting.svg')}
        alt=""
        style={{
          width: '20px',
          height: '20px',
          opacity: isActive ? 1 : 0.5,
          transform: isAsc ? 'scaleY(-1)' : 'none',
        }}
        className={styles.sortIcons}
        aria-hidden="true"
      />
    );
  };

  const renderFileNameCell = (doc: IDocument): React.ReactNode => {
    const matchKind = searchMatchKindByDocumentId?.get(doc.id);

    return (
      <div className={styles.fileNameCell}>
        {matchKind ? <SearchMatchBadge kind={matchKind} /> : null}
        <TooltipHost
          content={doc.fileName}
          className={styles.fileNameTextHost}
        >
          <span
            className={`${styles.cellTruncate} ${styles.cellLink} ${styles.fileNameLink}`}
            onClick={() => onDocumentClick(doc)}
            onKeyDown={(e) => {
              if (e.key === "Enter") onDocumentClick(doc);
            }}
            tabIndex={0}
            role="link"
            aria-label={`Open ${doc.fileName}`}
          >
            {truncateFileNameForDisplay(
              doc.fileName,
              matchKind
                ? FILE_NAME_DISPLAY_MAX_LENGTH_WITH_BADGE
                : FILE_NAME_DISPLAY_MAX_LENGTH,
            )}
          </span>
        </TooltipHost>
      </div>
    );
  };

  if (isLoading) {
    return (
      <div>
        {Array.from({ length: 5 }).map((_, i) => (
          <div key={i} className={`${styles.shimmer} ${styles.shimmerRow}`} />
        ))}
      </div>
    );
  }

  return (
    <div className={`${styles.dataTableScroll}${useEnhancedStyle ? ` ${styles.dataTableScrollEnhanced}` : ""}`}>
      <table
        className={styles.dataTable}
        role="grid"
        aria-label="Document results"
      >
        <thead>
          <tr>
            {(role === "contributor" || role === "admin") && (
              <th className={styles.leadingActionHeader} aria-label="Actions" />
            )}

            {visibleColumns.map((col) => (
              <th
                key={col.key}
                className={`${getColumnClassName(col.key)} ${col.sortField ? styles.sortable : ""}`.trim()}
                onClick={() => col.sortField && onSort(col.sortField)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && col.sortField) onSort(col.sortField);
                }}
                tabIndex={col.sortField ? 0 : undefined}
                role={col.sortField ? "columnheader button" : "columnheader"}
                aria-sort={
                  sortState && sortState.field === col.sortField
                    ? sortState.direction === "asc"
                      ? "ascending"
                      : "descending"
                    : undefined
                }
                aria-label={col.sortField ? `Sort by ${col.label}` : col.label}
              >
                {col.label}
                {renderSortIcon(col.sortField)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {documents.map((doc) => (
            <tr key={doc.id}>
              {(role === "contributor" || role === "admin") && (
                <td className={styles.leadingActionCell}>
                  {role === "contributor" &&
                    doc.status === "Current" &&
                    onArchiveReplaceClick && (
                      <TooltipHost content="Archive and Replace">
                        <span
                          className={styles.actionIcon}
                          onClick={(event) =>
                            onArchiveReplaceClick(doc, getAnchorPosition(event))
                          }
                          onKeyDown={(e) => {
                            if (e.key === "Enter") onArchiveReplaceClick(doc);
                          }}
                          tabIndex={0}
                          role="button"
                          aria-label={`Archive and replace ${doc.fileName}`}
                        >
                          <img src={require('../../assets/icons/archive-replace.svg')} alt="" style={{ width: '16px', height: '16px', display: 'block' }} aria-hidden="true" />
                        </span>
                      </TooltipHost>
                    )}

                  {role === "admin" && (
                    <div className={styles.rowActionsInline}>
                      <TooltipHost content="Edit metadata">
                        <button
                          className={styles.rowActionIconBtn}
                          onClick={() => onEditClick && onEditClick(doc)}
                          aria-label={`Edit ${doc.fileName}`}
                          type="button"
                        >
                          <img src={require('../../assets/icons/edit.svg')} alt="" style={{ width: '17px', height: '17px', display: 'block' }} aria-hidden="true" />
                        </button>
                      </TooltipHost>
                      {doc.status === "Current" && onArchiveReplaceClick && (
                        <TooltipHost content="Replace file and metadata">
                          <button
                            className={styles.rowActionIconBtn}
                            onClick={(event) =>
                              onArchiveReplaceClick(
                                doc,
                                getAnchorPosition(event),
                              )
                            }
                            aria-label={`Replace ${doc.fileName}`}
                            type="button"
                          >
                            <img src={require('../../assets/icons/archive-replace.svg')} alt="" style={{ width: '16px', height: '16px', display: 'block' }} aria-hidden="true" />
                          </button>
                        </TooltipHost>
                      )}
                      {doc.status === "Current" && (
                        <TooltipHost content="Archive document">
                          <button
                            className={styles.rowActionIconBtn}
                            onClick={() =>
                              onArchiveClick && onArchiveClick(doc)
                            }
                            aria-label={`Archive ${doc.fileName}`}
                            type="button"
                          >
                            <img src={require('../../assets/icons/archive.svg')} alt="" style={{ width: '16px', height: '16px', display: 'block' }} aria-hidden="true" />
                          </button>
                        </TooltipHost>
                      )}
                      {doc.status === "Archive" && (
                        <TooltipHost content="Re-activate document">
                          <button
                            className={styles.rowActionIconBtn}
                            onClick={() =>
                              onReActivateClick && onReActivateClick(doc)
                            }
                            aria-label={`Re-activate ${doc.fileName}`}
                            type="button"
                          >
                            <img src={require('../../assets/icons/re-activate.svg')} alt="" style={{ width: '16px', height: '16px', display: 'block' }} aria-hidden="true" />
                          </button>
                        </TooltipHost>
                      )}
                      <TooltipHost content="Delete document">
                        <button
                          className={styles.rowActionIconBtn}
                          onClick={() => onDeleteClick && onDeleteClick(doc)}
                          aria-label={`Delete ${doc.fileName}`}
                          type="button"
                          style={{ color: "#DC2626" }}
                        >
                          <img src={require('../../assets/icons/delete.svg')} alt="" style={{ width: '16px', height: '16px', display: 'block' }} aria-hidden="true" />
                        </button>
                      </TooltipHost>
                    </div>
                  )}
                </td>
              )}

              {visibleColumns.map((col) => (
                <td
                  key={col.key}
                  className={`${getColumnClassName(col.key)} ${col.key === "comments" ? styles.commentCell : ""}`.trim()}
                >
                  {col.key === "fileName"
                    ? renderFileNameCell(doc)
                    : col.key === "type" && useEnhancedStyle ? (
                    (() => {
                      const val = col.getValue(doc);
                      const chip = TYPE_CHIP_COLORS[val] ?? TYPE_CHIP_DEFAULT;
                      return (
                        <span
                          className={styles.typeChip}
                          style={{ backgroundColor: chip.bg, color: chip.color }}
                        >
                          {val}
                        </span>
                      );
                    })()
                  ) : col.key === "status" ? (
                    <span
                      className={
                        doc.status === "Current"
                          ? styles.statusActive
                          : styles.statusArchived
                      }
                    >
                      {doc.status === "Current" ? "Active" : "Archived"}
                    </span>
                  ) : col.key === "comments" ? (
                    <span className={styles.commentIcon}>
                      {(doc.comments || "").trim() ? (
                        <TooltipHost content={doc.comments.trim()}>
                          <img
                            src={require('../../assets/icons/comment.svg')}
                            alt=""
                            style={{ width: '16px', height: '16px', display: 'block' }}
                            aria-hidden="true"
                          />
                        </TooltipHost>
                      ) : null}
                    </span>
                  ) : col.truncate ? (
                    <TooltipHost content={col.getValue(doc)}>
                      <span className={styles.cellTruncate}>
                        {col.getValue(doc)}
                      </span>
                    </TooltipHost>
                  ) : (
                    col.getValue(doc)
                  )}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

