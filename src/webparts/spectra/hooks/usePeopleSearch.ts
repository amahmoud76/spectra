import { useState, useCallback, useRef, useEffect } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PeopleService, IPeopleSuggestion } from "../services/PeopleService";

const DEBOUNCE_MS = 350;
const MIN_QUERY_LENGTH = 2;

export interface IUsePeopleSearchResult {
  results: IPeopleSuggestion[];
  isSearching: boolean;
  search: (query: string) => void;
  clear: () => void;
}

export const usePeopleSearch = (
  context: WebPartContext,
  useMock: boolean,
): IUsePeopleSearchResult => {
  const [results, setResults] = useState<IPeopleSuggestion[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const serviceRef = useRef<PeopleService>(new PeopleService(context, useMock));

  // Clean up debounce timer on unmount
  useEffect(() => {
    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
    };
  }, []);

  const search = useCallback((query: string) => {
    if (timerRef.current) clearTimeout(timerRef.current);

    if (query.trim().length < MIN_QUERY_LENGTH) {
      setResults([]);
      setIsSearching(false);
      return;
    }

    setIsSearching(true);
    timerRef.current = setTimeout(async () => {
      try {
        const suggestions = await serviceRef.current.searchUsers(query.trim());
        setResults(suggestions);
      } catch {
        setResults([]);
      } finally {
        setIsSearching(false);
      }
    }, DEBOUNCE_MS);
  }, []);

  const clear = useCallback(() => {
    if (timerRef.current) clearTimeout(timerRef.current);
    setResults([]);
    setIsSearching(false);
  }, []);

  return { results, isSearching, search, clear };
};
