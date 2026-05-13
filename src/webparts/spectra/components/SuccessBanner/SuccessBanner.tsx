import * as React from "react";
import styles from "../SPECTRA.module.scss";

export interface ISuccessBannerProps {
  message: string;
  fileName?: string;
  onDismiss: () => void;
}

export const SuccessBanner: React.FC<ISuccessBannerProps> = ({
  message,
  fileName,
  onDismiss,
}) => {
  return (
    <div className={`${styles.banner} ${styles.bannerSuccess}`} role="status">
      <img src={require('../../assets/icons/check.svg')} alt="" className={`${styles.bannerIcon} ${styles.bannerIconSuccess}`} aria-hidden="true" style={{ width: '20px', height: '20px', display: 'block', flexShrink: 0 }} />
      <div className={styles.bannerContent}>
        <div className={styles.bannerMessage}>{message}</div>
        {fileName && (
          <div className={styles.bannerFileName}>
            File name applied: {fileName}
          </div>
        )}
      </div>
      <button
        className={styles.bannerDismiss}
        onClick={onDismiss}
        aria-label="Dismiss"
      >
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true"><path d="M1 1L11 11M11 1L1 11" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round"/></svg>
      </button>
    </div>
  );
};
