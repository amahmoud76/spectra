import * as React from "react";
import { format } from "date-fns";
import styles from "../SPECTRA.module.scss";
import abbvieLogo from "../../assets/abbvie.svg";
import {
  INACTIVITY_TIMEOUT_DEFAULT_MINUTES,
  INACTIVITY_TEST_MODE_MAX_MINUTES,
} from "../../config/config";

interface IFooterProps {
  inactivityTimeoutMinutes?: number;
}

const formatRemaining = (totalSeconds: number): string => {
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  const minutesText = minutes < 10 ? `0${minutes}` : String(minutes);
  const secondsText = seconds < 10 ? `0${seconds}` : String(seconds);
  return `${minutesText}:${secondsText}`;
};

export const Footer: React.FC<IFooterProps> = ({
  inactivityTimeoutMinutes = INACTIVITY_TIMEOUT_DEFAULT_MINUTES,
}) => {
  const [logoError, setLogoError] = React.useState(false);
  const [lastActivityAt, setLastActivityAt] = React.useState<number>(
    Date.now(),
  );
  const [now, setNow] = React.useState<number>(Date.now());
  const [isLocked, setIsLocked] = React.useState(false);

  const timeoutMs = Math.max(1, inactivityTimeoutMinutes) * 60 * 1000;
  const isTestMode =
    inactivityTimeoutMinutes <= INACTIVITY_TEST_MODE_MAX_MINUTES;
  const inactiveForMs = now - lastActivityAt;
  const remainingSeconds = Math.max(
    0,
    Math.ceil((timeoutMs - inactiveForMs) / 1000),
  );
  const isCountdownWarning = !isLocked && remainingSeconds <= 60;

  React.useEffect(() => {
    const onActivity = (): void => {
      if (isLocked) return;
      const timestamp = Date.now();
      setLastActivityAt(timestamp);
      setNow(timestamp);
    };

    const interval = window.setInterval(() => {
      setNow(Date.now());
    }, 1000);

    const events: Array<keyof WindowEventMap> = [
      "mousemove",
      "mousedown",
      "keydown",
      "scroll",
      "touchstart",
      "focus",
    ];

    events.forEach((eventName) => {
      window.addEventListener(eventName, onActivity);
    });

    return () => {
      window.clearInterval(interval);
      events.forEach((eventName) => {
        window.removeEventListener(eventName, onActivity);
      });
    };
  }, [isLocked]);

  React.useEffect(() => {
    if (isLocked) return;
    if (inactiveForMs >= timeoutMs) {
      setIsLocked(true);
    }
  }, [inactiveForMs, timeoutMs, isLocked]);

  return (
    <>
      <footer className={styles.footer}>
        {logoError ? (
          <div className={styles.footerBrand}>abbvie</div>
        ) : (
          <img
            src={abbvieLogo}
            alt="AbbVie"
            className={styles.footerLogo}
            onError={() => setLogoError(true)}
          />
        )}

        <div className={styles.footerMeta}>
          {isTestMode && (
            <span className={styles.testModeBadge}>Test Mode</span>
          )}
          {!isLocked && (
            <span
              className={`${styles.inactivityCountdown} ${isCountdownWarning ? styles.inactivityCountdownWarning : ""}`.trim()}
            >
              Session lock in {formatRemaining(remainingSeconds)}
            </span>
          )}
          <span>Copyright © {format(new Date(), "yyyy")} AbbVie Inc.</span>
          <span>v1.0</span>
        </div>
      </footer>

      {isLocked && (
        <div
          className={styles.sessionLockOverlay}
          role="alert"
          aria-live="assertive"
        >
          <div className={styles.sessionLockCard}>
            <h3>Session Locked</h3>
            <p>
              This session was locked after {inactivityTimeoutMinutes} minutes
              of inactivity.
            </p>
            <button
              className={styles.sessionUnlockButton}
              type="button"
              onClick={() => window.location.reload()}
            >
              Reload Session
            </button>
          </div>
        </div>
      )}
    </>
  );
};
