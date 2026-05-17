import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { SearchMatchKind } from "../../utils/filterHelper";
import styles from "../SPECTRA.module.scss";

export interface ISearchMatchBadgeProps {
  kind: SearchMatchKind;
}

export const SearchMatchBadge: React.FC<ISearchMatchBadgeProps> = ({ kind }) => {
  const isExact = kind === "exact";
  const label = isExact ? "Exact search match" : "Close search match";
  const letter = isExact ? "E" : "C";

  return (
    <TooltipHost content={label}>
      <span
        className={`${styles.searchMatchBadge} ${styles.searchMatchBadgeCompact} ${isExact ? styles.searchMatchBadgeExact : styles.searchMatchBadgeClose}`.trim()}
        aria-label={label}
      >
        {letter}
      </span>
    </TooltipHost>
  );
};
