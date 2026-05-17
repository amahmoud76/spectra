import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import styles from "../SPECTRA.module.scss";

export interface IViewFullLibraryButtonProps {
  onClick: () => void;
  isActive?: boolean;
}

export const ViewFullLibraryButton: React.FC<IViewFullLibraryButtonProps> = ({
  onClick,
  isActive = false,
}) => {
  return (
    <TooltipHost content="Browse all active documents in the library. Turn on Show Archived Documents to include archived items.">
      <button
        type="button"
        className={`${styles.btnSecondary} ${styles.viewFullLibraryBtn} ${isActive ? styles.viewFullLibraryBtnActive : ""}`.trim()}
        onClick={onClick}
        aria-label="View full document library"
        aria-pressed={isActive}
      >
        View Full Library
      </button>
    </TooltipHost>
  );
};
