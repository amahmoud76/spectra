import { useState, useCallback } from "react";
import { IFilterState, defaultFilterState } from "../interfaces/IFilterState";
import { countActiveFilters } from "../utils/filterHelper";

interface IUseFiltersResult {
  filters: IFilterState;
  setFilter: <K extends keyof IFilterState>(
    field: K,
    value: IFilterState[K],
  ) => void;
  clearAllFilters: () => void;
  clearFilter: (field: keyof IFilterState) => void;
  activeFilterCount: number;
  hasActiveFilters: boolean;
}

export const useFilters = (): IUseFiltersResult => {
  const [filters, setFilters] = useState<IFilterState>(defaultFilterState);

  const setFilter = useCallback(
    <K extends keyof IFilterState>(field: K, value: IFilterState[K]) => {
      setFilters((prev) => ({ ...prev, [field]: value }));
    },
    [],
  );

  const clearAllFilters = useCallback(() => {
    setFilters(defaultFilterState);
  }, []);

  const clearFilter = useCallback((field: keyof IFilterState) => {
    setFilters((prev) => ({ ...prev, [field]: defaultFilterState[field] }));
  }, []);

  const activeFilterCount = countActiveFilters(filters);

  return {
    filters,
    setFilter,
    clearAllFilters,
    clearFilter,
    activeFilterCount,
    hasActiveFilters: activeFilterCount > 0,
  };
};
