import * as React from "react";
import styles from "../SPECTRA.module.scss";

export interface IEmptyStateProps {
  type: "search-prompt" | "no-results" | "results-idle";
}

export const EmptyState: React.FC<IEmptyStateProps> = ({ type }) => {
  return (
    <div className={styles.emptyState} role="status">
      {type === "search-prompt" ? (
        <>
          <div className={styles.emptyStateIcon}>
            <i
              className="fa-solid fa-magnifying-glass"
              style={{ fontSize: 48, color: "#071D49" }}
            />
          </div>
          <div className={styles.emptyStateTitle}>Search for Documents</div>
          <div className={styles.emptyStateText}>
            Use the search bar or filters above to find documents.
          </div>
        </>
      ) : type === "no-results" ? (
        <>
          <div className={styles.emptyStateIcon}>
            <img src={require('../../assets/icons/no-results-found.svg')} alt="" style={{ width: '48px', height: '48px' }} />
          </div>
          <div className={styles.emptyStateTitle}>No Results Found</div>
          <div className={styles.emptyStateText}>
            Try checking spelling, using shorter terms, or adjusting filters.
          </div>
        </>
      ) : (
        <>
          <div className={styles.emptyStateIcon}>
            <i
              className="fa-solid fa-magnifying-glass"
              style={{ fontSize: 48, color: "#071D49" }}
            />
          </div>
          <div className={styles.emptyStateTitle}>Search Cleared</div>
          <div className={styles.emptyStateText}>
            Enter a new search term or apply filters to see documents.
          </div>
        </>
      )}
    </div>
  );
};
