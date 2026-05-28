import * as React from "react";
import styles from "../SPECTRA.module.scss";
import { captureAndLogError } from "../../services/errorLogService";

const MSAL_CLEAR_SESSION_KEY = "spectra_msal_clear_attempted";
const MSAL_CLEAR_COOLDOWN_MS = 30_000;

const wasRecentlyAttempted = (): boolean => {
  try {
    const stored = sessionStorage.getItem(MSAL_CLEAR_SESSION_KEY);
    if (!stored) return false;
    return Date.now() - parseInt(stored, 10) < MSAL_CLEAR_COOLDOWN_MS;
  } catch {
    return false;
  }
};

const clearMsalLocalStorageKeys = (): number => {
  try {
    const keys: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && /^msal\./.test(key)) keys.push(key);
    }
    keys.forEach((k) => localStorage.removeItem(k));
    return keys.length;
  } catch {
    return 0;
  }
};

interface IAuthWarningStripProps {
  onRetry: () => void;
  userEmail: string;
}

export const AuthWarningStrip: React.FC<IAuthWarningStripProps> = ({
  onRetry,
  userEmail,
}) => {
  const [dismissed, setDismissed] = React.useState(false);
  const showManualInstructions = wasRecentlyAttempted();

  if (dismissed) return null;

  const handleFixAndRetry = (): void => {
    const clearedCount = clearMsalLocalStorageKeys();

    void captureAndLogError(
      new Error("Auth recovery: MSAL localStorage cleared by user"),
      {
        component: "AuthWarningStrip",
        userAction: "Fix & Retry — MSAL cache clear",
        errorType: "Auth Recovery Attempt",
        featureFunction: "clearMsalAndRetry",
        userEmail,
        additionalContext: {
          clearedKeyCount: clearedCount,
          timestamp: new Date().toISOString(),
        },
      },
    );

    try {
      sessionStorage.setItem(MSAL_CLEAR_SESSION_KEY, Date.now().toString());
    } catch {
      // Ignore storage failures.
    }

    window.location.reload();
  };

  if (showManualInstructions) {
    return (
      <div className={styles.permStrip} role="alert">
        <img
          src={require("../../assets/icons/circle-exclamation.svg")}
          alt=""
          aria-hidden="true"
          className={styles.permStripIcon}
        />
        <span className={styles.permStripMessage}>
          Permissions still could not be verified. Try signing out of Microsoft 365 and back in, or clear your browser cache and reload.
        </span>
        <button
          className={styles.permStripDismiss}
          onClick={() => setDismissed(true)}
          aria-label="Dismiss"
        >
          <svg width="10" height="10" viewBox="0 0 10 10" fill="none" aria-hidden="true">
            <path d="M1 1L9 9M9 1L1 9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
          </svg>
        </button>
      </div>
    );
  }

  return (
    <div className={styles.permStrip} role="alert">
      <img
        src={require("../../assets/icons/circle-exclamation.svg")}
        alt=""
        aria-hidden="true"
        className={styles.permStripIcon}
      />
      <span className={styles.permStripMessage}>
        Permissions could not be verified — you have read-only access.
      </span>
      <button className={styles.permStripRetry} onClick={handleFixAndRetry}>
        Fix &amp; Retry
      </button>
      <button className={styles.permStripRetryLink} onClick={onRetry}>
        Retry without clearing
      </button>
      <button
        className={styles.permStripDismiss}
        onClick={() => setDismissed(true)}
        aria-label="Dismiss"
      >
        <svg width="10" height="10" viewBox="0 0 10 10" fill="none" aria-hidden="true">
          <path d="M1 1L9 9M9 1L1 9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
        </svg>
      </button>
    </div>
  );
};
