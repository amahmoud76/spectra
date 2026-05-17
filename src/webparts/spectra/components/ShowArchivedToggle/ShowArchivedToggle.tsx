import * as React from "react";
import styles from "../SPECTRA.module.scss";

export interface IShowArchivedToggleProps {
  checked: boolean;
  onChange: (checked: boolean) => void;
}

export const ShowArchivedToggle: React.FC<IShowArchivedToggleProps> = ({
  checked,
  onChange,
}) => {
  return (
    <div className={styles.archiveToggle}>
      <button
        className={styles.archiveToggleButton}
        type="button"
        role="switch"
        aria-checked={checked}
        aria-label="Show Archived Documents"
        onClick={() => onChange(!checked)}
      >
        <span
          className={`${styles.archiveToggleTrack} ${checked ? styles.archiveToggleTrackChecked : ""}`.trim()}
          aria-hidden="true"
        >
          <span className={styles.archiveToggleThumb}>
            {checked && (
              <img
                src={require("../../assets/icons/check-blue.svg")}
                alt=""
                className={styles.archiveToggleCheck}
                aria-hidden="true"
              />
            )}
          </span>
        </span>
        <span className={styles.archiveToggleLabel}>
          Show Archived Documents
        </span>
      </button>
    </div>
  );
};
