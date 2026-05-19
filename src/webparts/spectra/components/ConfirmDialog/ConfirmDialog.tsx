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
  onConfirm,
  onCancel,
  onSecondary,
}) => {
  // Close on Escape key
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent): void => {
      if (e.key === "Escape" && isOpen) onCancel();
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [isOpen, onCancel]);

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
      onClick={onCancel}
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
            >
              <i className="fa fa-xmark" />
            </button>
          </TooltipHost>
        </div>
        <div className={styles.modalBody}>{message}</div>
        <div className={styles.modalFooter}>
          <button className={styles.btnSecondary} onClick={onCancel}>
            {cancelLabel}
          </button>
          {secondaryLabel && onSecondary && (
            <button className={styles.btnSecondary} onClick={onSecondary}>
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
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};
