import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { ISortState, SortField } from "../../interfaces/ISortState";
import { SearchMatchKind } from "../../utils/filterHelper";
import {
  FILE_NAME_DISPLAY_MAX_LENGTH,
  FILE_NAME_DISPLAY_MAX_LENGTH_WITH_BADGE,
  FILE_NAME_DISPLAY_MAX_LENGTH_ENHANCED,
  FILE_NAME_DISPLAY_MAX_LENGTH_ENHANCED_WITH_BADGE,
  truncateFileNameForDisplay,
} from "../../utils/fileHelper";
import { DOCUMENT_TYPE_FULL_NAMES } from "../../config/config";
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
  onFavoriteToggle?: (doc: IDocument) => void;
  favoriteIds?: Set<string>;
  showFavorites?: boolean;
  isLoading: boolean;
  searchMatchKindByDocumentId?: Map<string, SearchMatchKind>;
  useEnhancedStyle?: boolean;
}

const FILE_TYPE_ICON: Record<string, string> = {
  pdf: require("../../assets/icons/file-pdf.svg"),
  docx: require("../../assets/icons/file-word.svg"),
  pptx: require("../../assets/icons/file-ppt.svg"),
  xlsx: require("../../assets/icons/file-excel.svg"),
};

const getFileExtension = (fileName: string): string => {
  const dot = fileName.lastIndexOf(".");
  return dot >= 0 ? fileName.substring(dot + 1).toLowerCase() : "";
};

const stripExtension = (fileName: string): string => {
  const dot = fileName.lastIndexOf(".");
  return dot >= 0 ? fileName.substring(0, dot) : fileName;
};

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

// ── Column width persistence ─────────────────────────────────
const COL_WIDTHS_KEY_PREFIX = "spectra_col_widths_v1_";

const loadColWidths = (role: string): Record<string, number> => {
  try {
    const raw = localStorage.getItem(COL_WIDTHS_KEY_PREFIX + role);
    return raw ? (JSON.parse(raw) as Record<string, number>) : {};
  } catch {
    return {};
  }
};

const saveColWidths = (role: string, widths: Record<string, number>): void => {
  try {
    localStorage.setItem(COL_WIDTHS_KEY_PREFIX + role, JSON.stringify(widths));
  } catch {
    /* quota exceeded — skip */
  }
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
    label: "Metadata Last Modified By",
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

const StarOutlineIcon: React.FC = () => (
  <svg
    width="16"
    height="16"
    viewBox="0 0 24 24"
    fill="none"
    aria-hidden="true"
  >
    <path
      d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"
      stroke="#6B7280"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const StarFilledIcon: React.FC = () => (
  <svg
    width="16"
    height="16"
    viewBox="0 0 24 24"
    fill="none"
    aria-hidden="true"
  >
    <path
      d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"
      fill="#0066f5"
      stroke="#0066f5"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

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
  onFavoriteToggle,
  favoriteIds,
  showFavorites,
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

  // ── Column resizing ──────────────────────────────────────────
  const [colWidths, setColWidths] = React.useState<Record<string, number>>(() =>
    loadColWidths(role),
  );

  const dragRef = React.useRef<{
    colKey: string;
    startX: number;
    startWidth: number;
    thEl: HTMLTableCellElement;
  } | null>(null);

  const onResizeMouseDown = React.useCallback(
    (e: React.MouseEvent, colKey: string): void => {
      e.preventDefault();
      e.stopPropagation();
      const thEl = (e.currentTarget as HTMLElement)
        .parentElement as HTMLTableCellElement;
      dragRef.current = {
        colKey,
        startX: e.clientX,
        startWidth: thEl.offsetWidth,
        thEl,
      };
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";

      const onMouseMove = (ev: MouseEvent): void => {
        if (!dragRef.current) return;
        const w = Math.max(
          50,
          dragRef.current.startWidth + ev.clientX - dragRef.current.startX,
        );
        dragRef.current.thEl.style.width = `${w}px`;
        dragRef.current.thEl.style.minWidth = `${w}px`;
      };

      const onMouseUp = (ev: MouseEvent): void => {
        if (!dragRef.current) return;
        const w = Math.max(
          50,
          dragRef.current.startWidth + ev.clientX - dragRef.current.startX,
        );
        const key = dragRef.current.colKey;
        dragRef.current = null;
        document.body.style.cursor = "";
        document.body.style.userSelect = "";
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
        setColWidths((prev) => {
          const next = { ...prev, [key]: w };
          saveColWidths(role, next);
          return next;
        });
      };

      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    },
    [role],
  );

  const onResizeReset = React.useCallback(
    (colKey: string): void => {
      setColWidths((prev) => {
        const next = { ...prev };
        delete next[colKey];
        saveColWidths(role, next);
        return next;
      });
    },
    [role],
  );

  // ── Mirrored top scrollbar ──────────────────────────────────
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const mirrorRef = React.useRef<HTMLDivElement>(null);
  const innerMirrorRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const scroll = scrollRef.current;
    const mirror = mirrorRef.current;
    if (!scroll || !mirror) return;
    const onScrollTable = (): void => {
      mirror.scrollLeft = scroll.scrollLeft;
    };
    const onScrollMirror = (): void => {
      scroll.scrollLeft = mirror.scrollLeft;
    };
    scroll.addEventListener("scroll", onScrollTable);
    mirror.addEventListener("scroll", onScrollMirror);
    return () => {
      scroll.removeEventListener("scroll", onScrollTable);
      mirror.removeEventListener("scroll", onScrollMirror);
    };
  }, []);

  React.useLayoutEffect(() => {
    const update = (): void => {
      if (scrollRef.current && innerMirrorRef.current) {
        innerMirrorRef.current.style.width = `${scrollRef.current.scrollWidth}px`;
      }
    };
    update();
    const ro = new ResizeObserver(update);
    if (scrollRef.current) ro.observe(scrollRef.current);
    return () => ro.disconnect();
  }, [documents, role, useEnhancedStyle]);

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
        src={require("../../assets/icons/sorting.svg")}
        alt=""
        style={{
          width: "20px",
          height: "20px",
          opacity: isActive ? 1 : 0.5,
          transform: isAsc ? "scaleY(-1)" : "none",
        }}
        className={styles.sortIcons}
        aria-hidden="true"
      />
    );
  };

  const renderFileNameCell = (doc: IDocument): React.ReactNode => {
    const matchKind = searchMatchKindByDocumentId?.get(doc.id);

    if (useEnhancedStyle) {
      const ext = getFileExtension(doc.fileName);
      const fileIconSrc = FILE_TYPE_ICON[ext];
      const displayName = stripExtension(doc.fileName);
      const maxLen = matchKind
        ? FILE_NAME_DISPLAY_MAX_LENGTH_ENHANCED_WITH_BADGE
        : FILE_NAME_DISPLAY_MAX_LENGTH_ENHANCED;

      return (
        <div className={styles.fileNameCell} style={{ gap: 0 }}>
          <div className={styles.fileNameBadgeArea}>
            {matchKind ? <SearchMatchBadge kind={matchKind} /> : null}
          </div>
          <div className={styles.fileNameDocContent}>
            {fileIconSrc && (
              <TooltipHost
                content={ext.toUpperCase()}
                styles={{ root: { display: "flex", alignItems: "center" } }}
              >
                <img
                  src={fileIconSrc}
                  alt={ext.toUpperCase()}
                  className={styles.fileTypeIcon}
                />
              </TooltipHost>
            )}
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
                {truncateFileNameForDisplay(displayName, maxLen)}
              </span>
            </TooltipHost>
          </div>
        </div>
      );
    }

    return (
      <div className={styles.fileNameCell}>
        {matchKind ? <SearchMatchBadge kind={matchKind} /> : null}
        <TooltipHost content={doc.fileName} className={styles.fileNameTextHost}>
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

  const mirrorDiv = (
    <div
      ref={mirrorRef}
      className={`${styleClassMap["scrollMirror"]}${useEnhancedStyle ? ` ${styleClassMap["scrollMirrorEnhanced"]}` : ""}`}
      aria-hidden="true"
    >
      <div
        ref={innerMirrorRef}
        className={styleClassMap["scrollMirrorInner"]}
      />
    </div>
  );

  const tableJsx = (
    <div
      ref={scrollRef}
      className={`${styles.dataTableScroll}${useEnhancedStyle ? ` ${styles.dataTableScrollEnhanced}` : ""}`}
    >
      <table
        className={styles.dataTable}
        role="grid"
        aria-label="Document results"
      >
        <thead>
          <tr>
            <th
              className={
                role === "admin"
                  ? styles.leadingActionHeader
                  : role === "contributor"
                    ? styles.leadingActionHeaderCompact
                    : styles.leadingActionHeaderNarrow
              }
              aria-label="Actions"
            />

            {visibleColumns.map((col) => (
              <th
                key={col.key}
                className={`${getColumnClassName(col.key)} ${col.sortField ? styles.sortable : ""}`.trim()}
                style={
                  colWidths[col.key]
                    ? {
                        width: colWidths[col.key],
                        minWidth: colWidths[col.key],
                      }
                    : undefined
                }
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
                <div
                  className={styleClassMap["colResizeHandle"]}
                  onMouseDown={(e) => onResizeMouseDown(e, col.key)}
                  onDoubleClick={(e) => {
                    e.stopPropagation();
                    onResizeReset(col.key);
                  }}
                  title="Drag to resize · Double-click to reset"
                  aria-hidden="true"
                />
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {documents.map((doc) => (
            <tr key={doc.id}>
              <td
                className={
                  role === "admin"
                    ? styles.leadingActionCell
                    : role === "contributor"
                      ? styles.leadingActionCellCompact
                      : styles.leadingActionCellNarrow
                }
              >
                <div className={styles.rowActionsInline}>
                  {/* Contributor-only: Archive & Replace */}
                  {role === "contributor" &&
                    doc.status === "Current" &&
                    onArchiveReplaceClick && (
                      <TooltipHost content="Archive and Replace">
                        <button
                          className={styles.rowActionIconBtn}
                          onClick={(event) =>
                            onArchiveReplaceClick(doc, getAnchorPosition(event))
                          }
                          aria-label={`Archive and replace ${doc.fileName}`}
                          type="button"
                        >
                          <img
                            src={require("../../assets/icons/archive-replace.svg")}
                            alt=""
                            style={{
                              width: "16px",
                              height: "16px",
                              display: "block",
                            }}
                            aria-hidden="true"
                          />
                        </button>
                      </TooltipHost>
                    )}

                  {/* Admin-only actions */}
                  {role === "admin" && (
                    <>
                      <TooltipHost content="Edit metadata">
                        <button
                          className={styles.rowActionIconBtn}
                          onClick={() => onEditClick && onEditClick(doc)}
                          aria-label={`Edit ${doc.fileName}`}
                          type="button"
                        >
                          <img
                            src={require("../../assets/icons/edit.svg")}
                            alt=""
                            style={{
                              width: "17px",
                              height: "17px",
                              display: "block",
                            }}
                            aria-hidden="true"
                          />
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
                            <img
                              src={require("../../assets/icons/archive-replace.svg")}
                              alt=""
                              style={{
                                width: "16px",
                                height: "16px",
                                display: "block",
                              }}
                              aria-hidden="true"
                            />
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
                            <img
                              src={require("../../assets/icons/archive.svg")}
                              alt=""
                              style={{
                                width: "16px",
                                height: "16px",
                                display: "block",
                              }}
                              aria-hidden="true"
                            />
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
                            <img
                              src={require("../../assets/icons/re-activate.svg")}
                              alt=""
                              style={{
                                width: "16px",
                                height: "16px",
                                display: "block",
                              }}
                              aria-hidden="true"
                            />
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
                          <img
                            src={require("../../assets/icons/delete.svg")}
                            alt=""
                            style={{
                              width: "16px",
                              height: "16px",
                              display: "block",
                            }}
                            aria-hidden="true"
                          />
                        </button>
                      </TooltipHost>
                    </>
                  )}

                  {/* Star / favourite — all roles, when feature enabled */}
                  {showFavorites && (
                    <TooltipHost
                      content={
                        favoriteIds?.has(doc.id)
                          ? "Remove from Favorites"
                          : "Add to Favorites"
                      }
                    >
                      <button
                        className={styles.rowActionIconBtn}
                        onClick={() =>
                          onFavoriteToggle && onFavoriteToggle(doc)
                        }
                        aria-label={
                          favoriteIds?.has(doc.id)
                            ? `Remove ${doc.fileName} from Favorites`
                            : `Add ${doc.fileName} to Favorites`
                        }
                        aria-pressed={favoriteIds?.has(doc.id)}
                        type="button"
                      >
                        {favoriteIds?.has(doc.id) ? (
                          <StarFilledIcon />
                        ) : (
                          <StarOutlineIcon />
                        )}
                      </button>
                    </TooltipHost>
                  )}
                </div>
              </td>

              {visibleColumns.map((col) => (
                <td
                  key={col.key}
                  className={`${getColumnClassName(col.key)} ${col.key === "comments" ? styles.commentCell : ""}`.trim()}
                  style={
                    colWidths[col.key]
                      ? {
                          width: colWidths[col.key],
                          minWidth: colWidths[col.key],
                        }
                      : undefined
                  }
                >
                  {col.key === "fileName" ? (
                    renderFileNameCell(doc)
                  ) : col.key === "type" && useEnhancedStyle ? (
                    (() => {
                      const val = col.getValue(doc);
                      const fullName = DOCUMENT_TYPE_FULL_NAMES[val];
                      return (
                        <TooltipHost
                          content={fullName ? `${val} – ${fullName}` : val}
                        >
                          <span className={styles.typeChip}>{val}</span>
                        </TooltipHost>
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
                            src={require("../../assets/icons/comment.svg")}
                            alt=""
                            style={{
                              width: "16px",
                              height: "16px",
                              display: "block",
                            }}
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
                  ) : col.key === "type" ? (
                    (() => {
                      const val = col.getValue(doc);
                      const fullName = DOCUMENT_TYPE_FULL_NAMES[val];
                      return fullName ? (
                        <TooltipHost content={`${val} – ${fullName}`}>
                          <span>{val}</span>
                        </TooltipHost>
                      ) : (
                        val
                      );
                    })()
                  ) : col.key === "paid" ? (
                    <div className={styles.cellPaidWrap}>
                      {col.getValue(doc)}
                    </div>
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

  return useEnhancedStyle ? (
    <>
      {mirrorDiv}
      <div className={styles.dataTableEnhancedOuter}>{tableJsx}</div>
    </>
  ) : (
    <>
      {mirrorDiv}
      {tableJsx}
    </>
  );
};
