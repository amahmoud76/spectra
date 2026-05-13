import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { DocumentViewer } from "../DocumentViewer/DocumentViewer";
import styles from "../SPECTRA.module.scss";

export interface IDocumentViewingPageProps {
  document: IDocument;
  role: EffectiveRole;
  siteUrl: string;
  onHome: () => void;
  onBack: () => void;
  onEditClick: () => void;
  onArchiveClick: () => void;
  onDeleteClick: () => void;
  onArchiveReplaceClick: () => void;
}

export const DocumentViewingPage: React.FC<IDocumentViewingPageProps> = ({
  document: doc,
  role,
  siteUrl,
  onHome,
  onBack,
  onEditClick,
  onArchiveClick,
  onDeleteClick,
  onArchiveReplaceClick,
}) => {
  return (
    <div className={styles.docViewPage}>
      {/* Breadcrumb */}
      <nav className={styles.breadcrumb} aria-label="Breadcrumb">
        <button
          className={styles.homeIconBtn}
          onClick={onHome}
          aria-label="Back to SPECTRA home"
          type="button"
        >
          <img src={require('../../assets/icons/home.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block' }} aria-hidden="true" />
        </button>
        <span
          className={styles.breadcrumbLink}
          onClick={onBack}
          onKeyDown={(e) => {
            if (e.key === "Enter") onBack();
          }}
          tabIndex={0}
          role="link"
        >
          SPECTRA Document Repository
        </span>
        <span className={styles.breadcrumbSeparator}>&gt;</span>
        <span className={styles.breadcrumbCurrent} title={doc.fileName}>
          {doc.fileName}
        </span>
      </nav>

      {/* Document name */}
      <div className={styles.docViewName}>
        <div>Document Name</div>
        <div className={styles.docViewNameValue}>{doc.fileName}</div>
      </div>

      {/* Role-specific action buttons */}
      <div className={styles.docViewActions}>
        <TooltipHost content="Return to document list">
          <button
            className={styles.docViewActionButton}
            onClick={onBack}
            aria-label="Return to document list"
            type="button"
          >
            Return
          </button>
        </TooltipHost>

        {/* Admin actions */}
        {role === "admin" && (
          <>
            <TooltipHost content="Edit metadata">
              <button
                className={styles.docViewActionButton}
                onClick={onEditClick}
                aria-label="Edit metadata"
              >
                <img src={require('../../assets/icons/edit.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block', marginRight: '4px' }} aria-hidden="true" /> Edit
              </button>
            </TooltipHost>
            {doc.status === "Current" && (
              <TooltipHost content="Archive document">
                <button
                  className={styles.docViewActionButton}
                  onClick={onArchiveClick}
                  aria-label="Archive document"
                >
                  <img src={require('../../assets/icons/archive.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block', marginRight: '4px' }} aria-hidden="true" /> Archive
                </button>
              </TooltipHost>
            )}
            <TooltipHost content="Delete document">
              <button
                className={styles.docViewActionButton}
                onClick={onDeleteClick}
                aria-label="Delete document"
              >
                <img src={require('../../assets/icons/delete.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block', marginRight: '4px' }} aria-hidden="true" /> Delete
              </button>
            </TooltipHost>
          </>
        )}

        {/* Contributor action */}
        {role === "contributor" && doc.status === "Current" && (
          <TooltipHost content="Archive and replace document">
            <button
              className={`${styles.docViewActionButton} ${styles.docViewActionButtonPrimary}`}
              onClick={onArchiveReplaceClick}
              aria-label="Archive and replace document"
            >
              <img src={require('../../assets/icons/archive-replace.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block', marginRight: '4px' }} aria-hidden="true" /> Archive and Replace
            </button>
          </TooltipHost>
        )}

        {/* Viewer — no action buttons */}
      </div>

      {/* Document Viewer — PDF.js for PDFs, embedview for Office files */}
      <DocumentViewer
        fileName={doc.fileName}
        fileUrl={doc.fileUrl}
        siteUrl={siteUrl}
      />
    </div>
  );
};
