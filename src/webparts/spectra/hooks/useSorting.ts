import { useState, useCallback } from "react";
import { IDocument } from "../interfaces/IDocument";
import { ISortState, SortField, SortDirection } from "../interfaces/ISortState";
import { sortDocuments } from "../utils/sortHelper";

interface IUseSortingResult {
  sortState: ISortState | null;
  userHasExplicitSort: boolean;
  toggleSort: (field: SortField) => void;
  clearSort: () => void;
  sortDocs: (documents: IDocument[]) => IDocument[];
}

/**
 * Toggle sort: Unsorted → Ascending → Descending → Unsorted
 * userHasExplicitSort is true only when the user has actively chosen a column sort.
 * It is false on initial load and after cycling back to unsorted.
 */
export const useSorting = (): IUseSortingResult => {
  const [sortState, setSortState] = useState<ISortState | null>({
    field: "uploadDate",
    direction: "desc",
  });
  const [userHasExplicitSort, setUserHasExplicitSort] = useState(false);

  const toggleSort = useCallback((field: SortField) => {
    setSortState((prev: ISortState | null): ISortState | null => {
      if (!prev || prev.field !== field) {
        setUserHasExplicitSort(true);
        return { field, direction: "asc" as SortDirection };
      }
      if (prev.direction === "asc") {
        setUserHasExplicitSort(true);
        return { field, direction: "desc" as SortDirection };
      }
      // Was desc → cycle back to match order
      setUserHasExplicitSort(false);
      return null;
    });
  }, []);

  const clearSort = useCallback(() => {
    setUserHasExplicitSort(false);
    setSortState({
      field: "uploadDate",
      direction: "desc",
    });
  }, []);

  const sortDocs = useCallback(
    (documents: IDocument[]) => sortDocuments(documents, sortState),
    [sortState],
  );

  return { sortState, userHasExplicitSort, toggleSort, clearSort, sortDocs };
};
