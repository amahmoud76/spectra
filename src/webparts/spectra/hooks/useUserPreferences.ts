import { useState, useCallback, useEffect, useRef } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  UserPreferencesService,
  IUserPreferences,
  ResultsView,
} from "../services/UserPreferencesService";

export interface IUseUserPreferencesResult {
  favorites: string[];
  recentlyViewed: string[];
  resultsView: ResultsView;
  isLoading: boolean;
  toggleFavorite: (docId: string) => Promise<{ added: boolean }>;
  addRecentlyViewed: (docId: string) => Promise<void>;
  setResultsView: (value: ResultsView) => Promise<void>;
}

export const useUserPreferences = (
  context: WebPartContext,
  userEmail: string,
  useMock: boolean,
  listName?: string,
): IUseUserPreferencesResult => {
  const [prefs, setPrefs] = useState<IUserPreferences>({
    listItemId: null,
    favorites: [],
    recentlyViewed: [],
    resultsView: "table",
  });
  const [isLoading, setIsLoading] = useState(true);
  const serviceRef = useRef<UserPreferencesService | null>(null);

  useEffect(() => {
    if (!userEmail) {
      setIsLoading(false);
      return;
    }
    const service = new UserPreferencesService(context, useMock, listName);
    serviceRef.current = service;

    service
      .getUserPreferences(userEmail)
      .then((p) => setPrefs(p))
      .catch(() => { /* silently fall back to empty prefs */ })
      .then(() => setIsLoading(false), () => setIsLoading(false));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userEmail]);

  const toggleFavorite = useCallback(
    async (docId: string): Promise<{ added: boolean }> => {
      const service = serviceRef.current;
      if (!service || !userEmail) return { added: false };

      const result = await service.toggleFavorite(userEmail, docId, prefs);
      setPrefs(result.prefs);
      return { added: result.added };
    },
    [prefs, userEmail],
  );

  const addRecentlyViewed = useCallback(
    async (docId: string): Promise<void> => {
      const service = serviceRef.current;
      if (!service || !userEmail) return;

      const newPrefs = await service.addRecentlyViewed(userEmail, docId, prefs);
      setPrefs(newPrefs);
    },
    [prefs, userEmail],
  );

  const setResultsView = useCallback(
    async (value: ResultsView): Promise<void> => {
      const service = serviceRef.current;
      if (!service || !userEmail) return;

      const newPrefs = await service.setResultsView(userEmail, value, prefs);
      setPrefs(newPrefs);
    },
    [prefs, userEmail],
  );

  return {
    favorites: prefs.favorites,
    recentlyViewed: prefs.recentlyViewed,
    resultsView: prefs.resultsView,
    isLoading,
    toggleFavorite,
    addRecentlyViewed,
    setResultsView,
  };
};
