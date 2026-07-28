import * as React from "react";
import styles from "../SPECTRA.module.scss";

export interface ISplashScreenProps {
  appName: string;
  subtitle?: string;
  statusLabel: string;
  progressPercent: number;
  isFading?: boolean;
  variant?: "classic" | "elegant";
  onAnimationComplete?: () => void;
}

const ELEGANT_STEPS = [
  { label: "Connecting to SharePoint", startAt: 0, doneAt: 38 },
  { label: "Loading user profile", startAt: 38, doneAt: 55 },
  { label: "Verifying permissions", startAt: 55, doneAt: 75 },
  { label: "Loading documents", startAt: 75, doneAt: 90 },
  { label: "Checking metadata", startAt: 90, doneAt: 98 },
  { label: "Building interface", startAt: 98, doneAt: 100 },
];

const STEP_MIN_MS = 600;

type StepState = "done" | "active" | "pending";

function getStepState(
  step: (typeof ELEGANT_STEPS)[0],
  progress: number,
): StepState {
  if (progress >= step.doneAt) return "done";
  if (progress >= step.startAt) return "active";
  return "pending";
}

const ElegantSplash: React.FC<ISplashScreenProps> = ({
  appName,
  subtitle,
  progressPercent,
  isFading = false,
  onAnimationComplete,
}) => {
  // Internal display progress paced at STEP_MIN_MS per step minimum.
  // It only advances past a step's doneAt when real progressPercent has
  // also reached that threshold, so the display never races ahead of reality.
  const [displayPct, setDisplayPct] = React.useState(0);
  const animDoneRef = React.useRef(false);

  React.useEffect(() => {
    if (animDoneRef.current) return;

    const activeStep = ELEGANT_STEPS.find(
      (s) => displayPct >= s.startAt && displayPct < s.doneAt,
    );

    if (!activeStep) {
      // All steps exhausted — animation complete
      if (!animDoneRef.current) {
        animDoneRef.current = true;
        onAnimationComplete?.();
      }
      return;
    }

    // Wait until real loading has actually reached this step's threshold
    if (progressPercent < activeStep.doneAt) return;

    const t = window.setTimeout(() => {
      setDisplayPct(activeStep.doneAt);
    }, STEP_MIN_MS);

    return () => window.clearTimeout(t);
  }, [displayPct, progressPercent, onAnimationComplete]);

  return (
    <div
      className={`${styles.splashElegant} ${isFading ? styles.splashElegantFading : ""}`}
      role="status"
      aria-live="polite"
    >
      <div className={styles.splashElegantDecor} aria-hidden="true" />

      <span className={styles.splashReleaseTag}>Release 1.3</span>

      <div className={styles.splashElegantInner}>
        <div className={styles.splashElegantBrand}>
          <img
            src={require("../../assets/spectra/Spectra-Stacked-LightMode.svg")}
            alt="SPECTRA logo"
            className={styles.splashElegantLogo}
          />
        </div>

        <div className={styles.splashElegantSteps}>
          {ELEGANT_STEPS.map((step) => {
            const state = getStepState(step, displayPct);
            return (
              <div
                key={step.label}
                className={[
                  styles.splashElegantStep,
                  state === "done"
                    ? styles.splashElegantStepDone
                    : state === "active"
                      ? styles.splashElegantStepActive
                      : styles.splashElegantStepPending,
                ].join(" ")}
              >
                <span className={styles.splashElegantStepIcon}>
                  {state === "done" && (
                    <svg
                      className={styles.splashElegantStepDoneIcon}
                      viewBox="0 0 16 16"
                      fill="none"
                      xmlns="http://www.w3.org/2000/svg"
                    >
                      <circle
                        cx="8"
                        cy="8"
                        r="8"
                        fill="rgba(92,219,110,0.15)"
                      />
                      <path
                        d="M4.5 8.5l2.5 2.5 4.5-5"
                        stroke="#5cdb6e"
                        strokeWidth="1.6"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                  )}
                  {state === "active" && (
                    <span className={styles.splashElegantStepSpinner} />
                  )}
                  {state === "pending" && (
                    <span className={styles.splashElegantStepCircle} />
                  )}
                </span>
                <span
                  className={[
                    styles.splashElegantStepLabel,
                    state === "active"
                      ? styles.splashElegantStepLabelActive
                      : "",
                  ].join(" ")}
                >
                  {step.label}
                </span>
              </div>
            );
          })}
        </div>

        <div className={styles.splashElegantProgress}>
          <div
            className={styles.splashElegantProgressFill}
            style={{ width: `${displayPct}%` }}
          />
        </div>
      </div>
    </div>
  );
};

const ClassicSplash: React.FC<ISplashScreenProps> = ({
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
            src={require("../../assets/spectra/Spectra-Stacked-LightMode.svg")}
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

export const SplashScreen: React.FC<ISplashScreenProps> = (props) => {
  if (props.variant === "classic") {
    // Classic variant has no animation pacing — signal complete immediately
    React.useEffect(() => {
      props.onAnimationComplete?.();
    }, []);
    return <ClassicSplash {...props} />;
  }
  return <ElegantSplash {...props} />;
};
