import { useState, useMemo, useCallback, useEffect } from "react";
import { IDocument } from "../interfaces/IDocument";
import { PAGE_SIZE_DEFAULT, PAGE_SIZE_OPTIONS } from "../config/config";

interface IUsePaginationResult {
  currentPage: number;
  pageSize: number;
  totalPages: number;
  paginatedDocuments: IDocument[];
  goToPage: (page: number) => void;
  goToFirstPage: () => void;
  goToLastPage: () => void;
  goToNextPage: () => void;
  goToPreviousPage: () => void;
  setPageSize: (size: number) => void;
  canGoNext: boolean;
  canGoPrevious: boolean;
  startIndex: number;
  endIndex: number;
  pageSizeOptions: number[];
}

export const usePagination = (
  documents: IDocument[],
  initialPageSize: number = PAGE_SIZE_DEFAULT,
): IUsePaginationResult => {
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSizeState] = useState(initialPageSize);

  const totalPages = useMemo(
    () => Math.max(1, Math.ceil(documents.length / pageSize)),
    [documents.length, pageSize],
  );

  // Reset to page 1 when documents or page size changes
  useEffect(() => {
    setCurrentPage(1);
  }, [documents.length, pageSize]);

  const paginatedDocuments = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return documents.slice(start, start + pageSize);
  }, [documents, currentPage, pageSize]);

  const goToPage = useCallback(
    (page: number) => {
      setCurrentPage(Math.max(1, Math.min(page, totalPages)));
    },
    [totalPages],
  );

  const goToFirstPage = useCallback(() => setCurrentPage(1), []);
  const goToLastPage = useCallback(
    () => setCurrentPage(totalPages),
    [totalPages],
  );
  const goToNextPage = useCallback(
    () => setCurrentPage((p) => Math.min(p + 1, totalPages)),
    [totalPages],
  );
  const goToPreviousPage = useCallback(
    () => setCurrentPage((p) => Math.max(p - 1, 1)),
    [],
  );

  const setPageSize = useCallback((size: number) => {
    setPageSizeState(size);
    setCurrentPage(1);
  }, []);

  const startIndex = (currentPage - 1) * pageSize + 1;
  const endIndex = Math.min(currentPage * pageSize, documents.length);

  return {
    currentPage,
    pageSize,
    totalPages,
    paginatedDocuments,
    goToPage,
    goToFirstPage,
    goToLastPage,
    goToNextPage,
    goToPreviousPage,
    setPageSize,
    canGoNext: currentPage < totalPages,
    canGoPrevious: currentPage > 1,
    startIndex,
    endIndex,
    pageSizeOptions: PAGE_SIZE_OPTIONS,
  };
};
