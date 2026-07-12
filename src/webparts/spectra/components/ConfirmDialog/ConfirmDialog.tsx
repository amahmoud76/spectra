import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import styles from "../SPECTRA.module.scss";

export interface IConfirmDialogProps {
  isOpen: boolean;
  title: string;
  message: React.ReactNode;
  confirmLabel: string;
  cancelLabel?: string;
  secondaryLabel?: string;
  isDestructive?: boolean;
  variant?: "default" | "compact";
  anchorPosition?: { top: number; left: number };
  /** Optional inline error shown inside the dialog (visible above the modal overlay). */
  errorMessage?: string;
  /** When true, the confirm/secondary actions are disabled (e.g. while an async action is in progress). */
  confirmDisabled?: boolean;
  /** When true, the dialog cannot be dismissed (Cancel, close, overlay click, Escape are all blocked). */
  disableDismiss?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
  onSecondary?: () => void;
}

export const ConfirmDialog: React.FC<IConfirmDialogProps> = ({
  isOpen,
  title,
  message,
  confirmLabel,
  cancelLabel = "Cancel",
  secondaryLabel,
  isDestructive = false,
  variant = "default",
  anchorPosition,
  errorMessage,
  confirmDisabled = false,
  disableDismiss = false,
  onConfirm,
  onCancel,
  onSecondary,
}) => {
  // Close on Escape key (unless dismissal is locked while processing)
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent): void => {
      if (e.key === "Escape" && isOpen && !disableDismiss) onCancel();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [isOpen, onCancel, disableDismiss]);

  if (!isOpen) return null;

  const compactModalStyle =
    variant === "compact" && anchorPosition
      ? {
          top: Math.min(anchorPosition.top, window.innerHeight - 280),
          left: Math.min(
            Math.max(12, anchorPosition.left),
            window.innerWidth - 440,
          ),
        }
      : undefined;

  return (
    <div
      className={`${styles.modalOverlay} ${variant === "compact" ? styles.modalOverlayCompact : ""}`}
      onClick={disableDismiss ? undefined : onCancel}
      role="dialog"
      aria-modal="true"
      aria-labelledby="confirm-dialog-title"
    >
      <div
        className={`${styles.modal} ${variant === "compact" ? styles.modalCompact : ""}`}
        onClick={(e) => e.stopPropagation()}
        style={compactModalStyle}
      >
        <div className={styles.modalHeader}>
          <h2 id="confirm-dialog-title" className={styles.modalTitle}>
            {title}
          </h2>
          <TooltipHost content="Close">
            <button
              className={styles.panelClose}
              onClick={onCancel}
              aria-label="Close dialog"
              disabled={disableDismiss}
            >
              <i className="fa fa-xmark" />
            </button>
          </TooltipHost>
        </div>
        <div className={styles.modalBody}>{message}</div>
        {errorMessage && (
          <div
            role="alert"
            style={{
              margin: "0 24px 12px",
              padding: "10px 12px",
              borderRadius: "6px",
              background: "#fdecea",
              border: "1px solid #f5c2c0",
              color: "#b42318",
              fontSize: "13px",
            }}
          >
            {errorMessage}
          </div>
        )}
        <div className={styles.modalFooter}>
          <button
            className={styles.btnSecondary}
            onClick={onCancel}
            disabled={disableDismiss}
          >
            {cancelLabel}
          </button>
          {secondaryLabel && onSecondary && (
            <button
              className={styles.btnSecondary}
              onClick={onSecondary}
              disabled={confirmDisabled}
            >
              {secondaryLabel}
            </button>
          )}
          <button
            className={
              isDestructive
                ? `${styles.btnPrimary} ${styles.btnDelete}`
                : styles.btnPrimary
            }
            onClick={onConfirm}
            disabled={confirmDisabled}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};
