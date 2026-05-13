import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import styles from "../SPECTRA.module.scss";

export interface ISearchBarProps {
  value: string;
  onChange: (value: string) => void;
  onSearch: () => void;
  onClear?: () => void;
  placeholder?: string;
  className?: string;
  isError?: boolean;
}

export const SearchBar: React.FC<ISearchBarProps> = ({
  value,
  onChange,
  onSearch,
  onClear,
  placeholder = "Search by document name, asset, type, or therapeutic area",
  className,
  isError = false,
}) => {
  const handleKeyDown = React.useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === "Enter") {
        e.preventDefault();
        onSearch();
      }
    },
    [onSearch],
  );

  const handleClear = React.useCallback(() => {
    if (onClear) {
      onClear();
      return;
    }

    onChange("");
  }, [onChange, onClear]);

  return (
    <div
      className={`${styles.searchBar} ${isError ? styles.searchBarError : ""} ${className || ""}`.trim()}
    >
      <TooltipHost content="Search">
        <button
          className={styles.searchBarAction}
          onClick={onSearch}
          aria-label="Search documents"
          type="button"
        >
          <img src={require('../../assets/icons/search.svg')} alt="" className={styles.searchBarIcon} aria-hidden="true" style={{ width: '16px', height: '16px', display: 'block' }} />
        </button>
      </TooltipHost>
      <input
        type="text"
        className={styles.searchBarInput}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={handleKeyDown}
        placeholder={placeholder}
        aria-label="Search documents"
      />
      {value && (
        <TooltipHost content="Clear search">
          <button
            className={styles.searchBarClear}
            onClick={handleClear}
            aria-label="Clear search"
            type="button"
          >
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
              <path d="M1 1L11 11M11 1L1 11" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round"/>
            </svg>
          </button>
        </TooltipHost>
      )}
    </div>
  );
};
