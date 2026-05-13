import * as React from "react";
import styles from "../SPECTRA.module.scss";

export interface IErrorBannerProps {
  message: string;
  onDismiss: () => void;
}

export const ErrorBanner: React.FC<IErrorBannerProps> = ({
  message,
  onDismiss,
}) => {
  return (
    <div className={`${styles.banner} ${styles.bannerError}`} role="alert">
      <img
        src={require("../../assets/icons/ban.svg")}
        alt=""
        style={{ width: "20px", height: "20px", display: "block", flexShrink: 0 }}
        className={`${styles.bannerIcon} ${styles.bannerIconError}`}
      />
      <div className={styles.bannerContent}>
        <div className={styles.bannerMessage}>{message}</div>
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
