import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { SearchMatchKind } from "../../utils/filterHelper";
import { SearchMatchBadge } from "../SearchMatchBadge/SearchMatchBadge";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { parseISO, format, isValid } from "date-fns";
import { DOCUMENT_TYPE_FULL_NAMES } from "../../config/config";
import { formatAssetWithSynonyms } from "../../utils/assetSearchAliasHelper";
import styles from "./TilesView.module.scss";
import parentStyles from "../SPECTRA.module.scss";

export interface ITilesViewProps {
  documents: IDocument[];
  role: EffectiveRole;
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
}

const FILE_TYPE_ICON: Record<string, string> = {
  pdf: require("../../assets/icons/file-pdf.svg"),
  docx: require("../../assets/icons/file-word.svg"),
  pptx: require("../../assets/icons/file-ppt.svg"),
  xlsx: require("../../assets/icons/file-excel.svg"),
};

const formatDate = (iso: string): string => {
  if (!iso) return "";
  const date = parseISO(iso);
  if (!isValid(date)) return iso;
  return format(date, "MMM/dd/yyyy");
};

const getFileExtension = (fileName: string): string => {
  const dot = fileName.lastIndexOf(".");
  return dot >= 0 ? fileName.substring(dot + 1).toLowerCase() : "";
};

const KebabIcon: React.FC = () => (
  <svg
    width="18"
    height="18"
    viewBox="0 0 24 24"
    fill="none"
    aria-hidden="true"
  >
    <circle cx="12" cy="5" r="2" fill="#0066F5" />
    <circle cx="12" cy="12" r="2" fill="#0066F5" />
    <circle cx="12" cy="19" r="2" fill="#0066F5" />
  </svg>
);

export const TilesView: React.FC<ITilesViewProps> = ({
  documents,
  role,
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
}) => {
  const [expandedIds, setExpandedIds] = React.useState<Set<string>>(
    () => new Set(),
  );

  // ── Admin actions overflow menu (single menu open at a time) ──
  const [openMenuDocId, setOpenMenuDocId] = React.useState<string | null>(null);
  const openMenuRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    if (!openMenuDocId) return;
    const handler = (e: MouseEvent): void => {
      if (
        openMenuRef.current &&
        !openMenuRef.current.contains(e.target as Node)
      ) {
        setOpenMenuDocId(null);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [openMenuDocId]);

  const toggleExpanded = React.useCallback((id: string) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }, []);

  const getAnchorPosition = React.useCallback(
    (event: React.MouseEvent<HTMLElement>) => {
      const rect = event.currentTarget.getBoundingClientRect();
      return { top: rect.bottom + 8, left: rect.left };
    },
    [],
  );

  if (isLoading) {
    return (
      <div className={styles.loadingState} role="status" aria-live="polite">
        Loading…
      </div>
    );
  }

  if (documents.length === 0) {
    return null;
  }

  const renderField = (
    label: string,
    value: React.ReactNode,
    wrapClassName?: string,
  ): JSX.Element => (
    <div
      className={
        wrapClassName
          ? `${styles.tileField} ${wrapClassName}`
          : styles.tileField
      }
    >
      <div className={styles.tileFieldLabel}>{label}</div>
      <div className={styles.tileFieldValue}>{value || "—"}</div>
    </div>
  );

  // ── Role-based layout ────────────────────────────────────────
  // Action column width: admin has a kebab menu + star, contributor 2 icons, viewer 1
  const actionColWidth =
    role === "admin" ? "72px" : role === "contributor" ? "64px" : "32px";
  const rowGridTemplate = `36px ${actionColWidth} 1fr`;

  // Body / expanded grid: admin includes Status column (9 cols), others use 8
  const tileBodyGrid =
    role === "admin"
      ? "minmax(220px, 2fr) minmax(90px, 1fr) minmax(90px, 1fr) minmax(80px, 1fr) minmax(90px, 1fr) minmax(100px, 1.2fr) minmax(90px, 1fr) minmax(100px, 1fr) minmax(80px, 0.9fr)"
      : "minmax(220px, 2fr) minmax(90px, 1fr) minmax(90px, 1fr) minmax(80px, 1fr) minmax(90px, 1fr) minmax(100px, 1.2fr) minmax(90px, 1fr) minmax(100px, 1fr)";

  return (
    <div className={styles.tilesContainer}>
      {documents.map((doc) => {
        const isExpanded = expandedIds.has(doc.id);
        const isFavorite = favoriteIds?.has(doc.id) ?? false;
        const matchKind = searchMatchKindByDocumentId?.get(doc.id);
        const ext = getFileExtension(doc.fileName);
        const fileIconSrc = FILE_TYPE_ICON[ext];
        const fullTypeName = DOCUMENT_TYPE_FULL_NAMES[doc.documentType];

        return (
          <div key={doc.id} className={styles.tile}>
            <div
              className={styles.tileRow}
              style={{ gridTemplateColumns: rowGridTemplate }}
            >
              <button
                className={`${styles.expandBtn} ${isExpanded ? styles.expandBtnOpen : ""}`}
                onClick={() => toggleExpanded(doc.id)}
                aria-expanded={isExpanded}
                aria-label={isExpanded ? "Collapse details" : "Expand details"}
                type="button"
              >
                <img
                  src={require("../../assets/icons/expand.svg")}
                  alt=""
                  aria-hidden="true"
                  className={styles.expandIcon}
                />
              </button>

              <div className={styles.tileActions}>
                {role === "admin" && (
                  <div
                    className={parentStyles.rowActionsMenuWrap}
                    ref={openMenuDocId === doc.id ? openMenuRef : undefined}
                  >
                    <button
                      className={parentStyles.rowActionsMenuTrigger}
                      onClick={() =>
                        setOpenMenuDocId(
                          openMenuDocId === doc.id ? null : doc.id,
                        )
                      }
                      aria-label={`Actions for ${doc.fileName}`}
                      aria-haspopup="menu"
                      aria-expanded={openMenuDocId === doc.id}
                      type="button"
                    >
                      <KebabIcon />
                    </button>

                    {openMenuDocId === doc.id && (
                      <div
                        className={parentStyles.rowActionsMenu}
                        role="menu"
                        aria-label={`Actions for ${doc.fileName}`}
                      >
                        <button
                          className={parentStyles.rowActionsMenuItem}
                          role="menuitem"
                          type="button"
                          onClick={() => {
                            setOpenMenuDocId(null);
                            onEditClick && onEditClick(doc);
                          }}
                        >
                          <span>Edit</span>
                          <img
                            src={require("../../assets/icons/edit.svg")}
                            alt=""
                            className={parentStyles.rowActionsMenuItemIcon}
                            aria-hidden="true"
                          />
                        </button>

                        {doc.status === "Current" && onArchiveReplaceClick && (
                          <button
                            className={parentStyles.rowActionsMenuItem}
                            role="menuitem"
                            type="button"
                            onClick={(event) => {
                              setOpenMenuDocId(null);
                              onArchiveReplaceClick(
                                doc,
                                getAnchorPosition(event),
                              );
                            }}
                          >
                            <span>Replace</span>
                            <img
                              src={require("../../assets/icons/archive-replace.svg")}
                              alt=""
                              className={parentStyles.rowActionsMenuItemIcon}
                              aria-hidden="true"
                            />
                          </button>
                        )}

                        {doc.status === "Current" && (
                          <button
                            className={parentStyles.rowActionsMenuItem}
                            role="menuitem"
                            type="button"
                            onClick={() => {
                              setOpenMenuDocId(null);
                              onArchiveClick && onArchiveClick(doc);
                            }}
                          >
                            <span>Archive</span>
                            <img
                              src={require("../../assets/icons/archive.svg")}
                              alt=""
                              className={parentStyles.rowActionsMenuItemIcon}
                              aria-hidden="true"
                            />
                          </button>
                        )}

                        {doc.status === "Archive" && (
                          <button
                            className={parentStyles.rowActionsMenuItem}
                            role="menuitem"
                            type="button"
                            onClick={() => {
                              setOpenMenuDocId(null);
                              onReActivateClick && onReActivateClick(doc);
                            }}
                          >
                            <span>Re-activate</span>
                            <img
                              src={require("../../assets/icons/re-activate.svg")}
                              alt=""
                              className={parentStyles.rowActionsMenuItemIcon}
                              aria-hidden="true"
                            />
                          </button>
                        )}

                        <button
                          className={parentStyles.rowActionsMenuItem}
                          role="menuitem"
                          type="button"
                          onClick={() => {
                            setOpenMenuDocId(null);
                            onDeleteClick && onDeleteClick(doc);
                          }}
                        >
                          <span>Delete</span>
                          <img
                            src={require("../../assets/icons/delete.svg")}
                            alt=""
                            className={parentStyles.rowActionsMenuItemIcon}
                            aria-hidden="true"
                          />
                        </button>
                      </div>
                    )}
                  </div>
                )}
                {role === "contributor" &&
                  doc.status === "Current" &&
                  onArchiveReplaceClick && (
                    <TooltipHost content="Archive and Replace">
                      <button
                        className={parentStyles.rowActionIconBtn}
                        onClick={(event) =>
                          onArchiveReplaceClick(doc, getAnchorPosition(event))
                        }
                        aria-label={`Archive and replace ${doc.fileName}`}
                        type="button"
                      >
                        <img
                          src={require("../../assets/icons/archive-replace.svg")}
                          alt=""
                          aria-hidden="true"
                          style={{ width: 16, height: 16, display: "block" }}
                        />
                      </button>
                    </TooltipHost>
                  )}
                {showFavorites && (
                  <TooltipHost
                    content={
                      isFavorite ? "Remove from Favorites" : "Add to Favorites"
                    }
                  >
                    <button
                      className={parentStyles.rowActionIconBtn}
                      onClick={() => onFavoriteToggle && onFavoriteToggle(doc)}
                      aria-label={
                        isFavorite
                          ? `Remove ${doc.fileName} from Favorites`
                          : `Add ${doc.fileName} to Favorites`
                      }
                      aria-pressed={isFavorite}
                      type="button"
                    >
                      <img
                        src={
                          isFavorite
                            ? require("../../assets/icons/star-full.svg")
                            : require("../../assets/icons/star-empty.svg")
                        }
                        alt=""
                        aria-hidden="true"
                        style={{ width: 17, height: 17, display: "block" }}
                      />
                    </button>
                  </TooltipHost>
                )}
              </div>

              <div
                className={styles.tileBody}
                style={{ gridTemplateColumns: tileBodyGrid }}
              >
                {renderField(
                  "Document Name",
                  <div className={styles.fileNameWrap}>
                    {matchKind ? <SearchMatchBadge kind={matchKind} /> : null}
                    {fileIconSrc && (
                      <img
                        src={fileIconSrc}
                        alt={ext.toUpperCase()}
                        className={styles.fileTypeIcon}
                      />
                    )}
                    <TooltipHost
                      content={doc.fileName}
                      className={parentStyles.fileNameTextHost}
                    >
                      <span
                        className={styles.fileNameLink}
                        onClick={() => onDocumentClick(doc)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") onDocumentClick(doc);
                        }}
                        tabIndex={0}
                        role="link"
                        aria-label={`Open ${doc.fileName}`}
                      >
                        {doc.fileName}
                      </span>
                    </TooltipHost>
                  </div>,
                )}
                {renderField(
                  "Asset",
                  (() => {
                    const assetLabel = formatAssetWithSynonyms(
                      doc.asset,
                      doc.assetSearchTokens,
                    );
                    return assetLabel ? (
                      <TooltipHost content={assetLabel}>
                        <span>{assetLabel}</span>
                      </TooltipHost>
                    ) : (
                      ""
                    );
                  })(),
                )}
                {renderField(
                  "Type",
                  fullTypeName ? (
                    <TooltipHost content={fullTypeName}>
                      <span>{doc.documentType}</span>
                    </TooltipHost>
                  ) : (
                    doc.documentType
                  ),
                )}
                {renderField("TA", doc.therapeuticArea.join("; "))}
                {renderField("Indication", doc.indication.join("; "))}
                {renderField("PAID", doc.paid.join("; "))}
                {renderField("DAS", doc.diseaseArea.join("; "))}
                {renderField("Eff Date", formatDate(doc.effectiveDate))}
                {role === "admin" &&
                  renderField(
                    "Status",
                    <span
                      className={
                        doc.status === "Current"
                          ? styles.statusActive
                          : styles.statusArchive
                      }
                    >
                      {doc.status === "Current" ? "Active" : "Archived"}
                    </span>,
                  )}
              </div>
            </div>

            <div
              className={`${styles.tileExpandedWrap} ${isExpanded ? styles.tileExpandedWrapOpen : ""}`}
              aria-hidden={!isExpanded}
            >
              <div className={styles.tileExpandedRow}>
                <div />
                <div className={styles.tileExpanded}>
                  {role === "admin" &&
                    renderField("Upload Date", formatDate(doc.uploadDate))}
                  {renderField("Created by", doc.createdBy)}
                  {renderField("Metadata Last Modified by", doc.modifiedBy)}
                  {renderField(
                    "Comment",
                    (doc.comments || "").trim() || "—",
                    role === "admin"
                      ? styles.tileFieldComment
                      : styles.tileFieldCommentShort,
                  )}
                </div>
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
};
