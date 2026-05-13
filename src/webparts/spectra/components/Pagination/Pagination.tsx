import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import styles from "../SPECTRA.module.scss";

export interface IPaginationProps {
  currentPage: number;
  totalPages: number;
  pageSize: number;
  pageSizeOptions: number[];
  startIndex: number;
  endIndex: number;
  totalCount: number;
  canGoNext: boolean;
  canGoPrevious: boolean;
  onPageChange: (page: number) => void;
  onPageSizeChange: (size: number) => void;
  onFirst: () => void;
  onPrevious: () => void;
  onNext: () => void;
  onLast: () => void;
}

export const Pagination: React.FC<IPaginationProps> = ({
  currentPage,
  totalPages,
  pageSize,
  pageSizeOptions,
  startIndex,
  endIndex,
  totalCount,
  canGoNext,
  canGoPrevious,
  onPageChange,
  onPageSizeChange,
  onFirst,
  onPrevious,
  onNext,
  onLast,
}) => {
  // Generate visible page numbers (max 7)
  const getPageNumbers = (): number[] => {
    const pages: number[] = [];
    const maxVisible = 7;

    if (totalPages <= maxVisible) {
      for (let i = 1; i <= totalPages; i++) pages.push(i);
    } else {
      const start = Math.max(1, currentPage - 3);
      const end = Math.min(totalPages, start + maxVisible - 1);
      for (let i = start; i <= end; i++) pages.push(i);
    }

    return pages;
  };

  return (
    <div
      className={styles.pagination}
      role="navigation"
      aria-label="Pagination"
    >
      {/* Page size selector */}
      <div className={styles.paginationSizes}>
        <span>View:</span>
        {pageSizeOptions.map((size) => (
          <TooltipHost key={size} content={`Show ${size} per page`}>
            <button
              className={`${styles.paginationSizeBtn} ${size === pageSize ? styles.paginationSizeBtnActive : ""}`}
              onClick={() => onPageSizeChange(size)}
              aria-label={`Show ${size} items per page`}
              aria-pressed={size === pageSize}
            >
              {size}
            </button>
          </TooltipHost>
        ))}
      </div>

      {/* Result count */}
      <span>
        {startIndex} - {endIndex} of {totalCount}
      </span>

      {/* Page navigation */}
      <div className={styles.paginationNav}>
        <TooltipHost content="First page">
          <button
            className={styles.paginationNavBtn}
            onClick={onFirst}
            disabled={!canGoPrevious}
            aria-label="First page"
          >
            «
          </button>
        </TooltipHost>
        <TooltipHost content="Previous page">
          <button
            className={styles.paginationNavBtn}
            onClick={onPrevious}
            disabled={!canGoPrevious}
            aria-label="Previous page"
          >
            ‹
          </button>
        </TooltipHost>

        {getPageNumbers().map((page) => (
          <button
            key={page}
            className={`${styles.paginationNavBtn} ${page === currentPage ? styles.paginationNavBtnActive : ""}`}
            onClick={() => onPageChange(page)}
            aria-label={`Page ${page}`}
            aria-current={page === currentPage ? "page" : undefined}
          >
            {page}
          </button>
        ))}

        <TooltipHost content="Next page">
          <button
            className={styles.paginationNavBtn}
            onClick={onNext}
            disabled={!canGoNext}
            aria-label="Next page"
          >
            ›
          </button>
        </TooltipHost>
        <TooltipHost content="Last page">
          <button
            className={styles.paginationNavBtn}
            onClick={onLast}
            disabled={!canGoNext}
            aria-label="Last page"
          >
            »
          </button>
        </TooltipHost>
      </div>
    </div>
  );
};
