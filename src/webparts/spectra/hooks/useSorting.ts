import { useState, useCallback } from "react";
import { IDocument } from "../interfaces/IDocument";
import { ISortState, SortField, SortDirection } from "../interfaces/ISortState";
import { sortDocuments } from "../utils/sortHelper";

interface IUseSortingResult {
  sortState: ISortState | null;
  toggleSort: (field: SortField) => void;
  clearSort: () => void;
  sortDocs: (documents: IDocument[]) => IDocument[];
}

/**
 * Toggle sort: Unsorted → Ascending → Descending → Unsorted
 */
export const useSorting = (): IUseSortingResult => {
  const [sortState, setSortState] = useState<ISortState | null>({
    field: "uploadDate",
    direction: "desc",
  });

  const toggleSort = useCallback((field: SortField) => {
    setSortState((prev: ISortState | null): ISortState | null => {
      if (!prev || prev.field !== field) {
        return { field, direction: "asc" as SortDirection };
      }
      if (prev.direction === "asc") {
        return { field, direction: "desc" as SortDirection };
      }
      // Was desc → clear sort
      return null;
    });
  }, []);

  const clearSort = useCallback(() => {
    setSortState({
      field: "uploadDate",
      direction: "desc",
    });
  }, []);

  const sortDocs = useCallback(
    (documents: IDocument[]) => sortDocuments(documents, sortState),
    [sortState],
  );

  return { sortState, toggleSort, clearSort, sortDocs };
};
