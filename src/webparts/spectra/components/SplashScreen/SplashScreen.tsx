import * as React from "react";
import styles from "../SPECTRA.module.scss";

export interface ISplashScreenProps {
  appName: string;
  subtitle?: string;
  statusLabel: string;
  progressPercent: number;
  isFading?: boolean;
}

export const SplashScreen: React.FC<ISplashScreenProps> = ({
  appName,
  subtitle,
  statusLabel,
  progressPercent,
  isFading = false,
}) => {
  const clampedProgress = Math.min(100, Math.max(0, progressPercent));

  return (
    <div
      className={`${styles.splashScreen} ${isFading ? styles.splashScreenFading : ""}`}
      role="status"
      aria-live="polite"
    >
      <div className={styles.splashCard}>
        <div className={styles.splashBrandRow}>
          <img
            src={require("../../assets/icons/rainbow-half.svg")}
            alt="SPECTRA logo"
            className={styles.splashLogo}
          />
          <h1 className={styles.splashTitle}>{appName}</h1>
        </div>
        {subtitle && <p className={styles.splashSubtitle}>{subtitle}</p>}
        <p className={styles.splashStatus}>{statusLabel}</p>

        <div className={styles.splashProgressWrap}>
          <div className={styles.splashProgressTrack}>
            <div
              className={styles.splashProgressFill}
              style={{ width: `${clampedProgress}%` }}
            />
          </div>
        </div>
      </div>
    </div>
  );
};
