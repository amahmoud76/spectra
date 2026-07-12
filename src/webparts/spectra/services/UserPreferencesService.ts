import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export const DEFAULT_USER_PREFS_LIST_NAME = "SPECTRA_UserPreferences";
const MAX_RECENTLY_VIEWED = 10;
const MOCK_STORAGE_KEY = "spectra_user_prefs_mock";
// Per-user fallback key used when the user lacks Contribute on the SP list (e.g. View Only role)
const localFallbackKey = (email: string): string =>
  `spectra_user_prefs_local_${email.toLowerCase().replace(/[^a-z0-9]/g, "_")}`;

const loadLocalFallback = (email: string): Partial<IUserPreferences> => {
  try {
    const raw = localStorage.getItem(localFallbackKey(email));
    return raw ? (JSON.parse(raw) as Partial<IUserPreferences>) : {};
  } catch {
    return {};
  }
};

const saveLocalFallback = (email: string, prefs: IUserPreferences): void => {
  try {
    localStorage.setItem(localFallbackKey(email), JSON.stringify({
      favorites: prefs.favorites,
      recentlyViewed: prefs.recentlyViewed,
      resultsView: prefs.resultsView,
    }));
  } catch { /* ignore quota errors */ }
};

export interface IUserPreferences {
  listItemId: number | null;
  favorites: string[];
  recentlyViewed: string[];
  resultsView: ResultsView;
}

export type ResultsView = "table" | "tiles";

interface IRawPrefsItem {
  Id: number;
  Favorites: string | null;
  RecentlyViewed: string | null;
  ResultsView: string | null;
}

export class UserPreferencesService {
  private context: WebPartContext;
  private useMock: boolean;
  private listInitialized = false;
  private readonly listName: string;

  constructor(context: WebPartContext, useMock: boolean = false, listName: string = DEFAULT_USER_PREFS_LIST_NAME) {
    this.context = context;
    this.useMock = useMock;
    this.listName = listName;
  }

  private get siteUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  // ── Mock helpers ──────────────────────────────────────────────

  private mockLoad(): IUserPreferences {
    try {
      const raw = localStorage.getItem(MOCK_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as Partial<IUserPreferences>;
        return {
          listItemId: parsed.listItemId ?? null,
          favorites: parsed.favorites ?? [],
          recentlyViewed: parsed.recentlyViewed ?? [],
          resultsView: parsed.resultsView === "tiles" ? "tiles" : "table",
        };
      }
    } catch { /* ignore */ }
    return { listItemId: null, favorites: [], recentlyViewed: [], resultsView: "table" };
  }

  private mockSave(prefs: IUserPreferences): void {
    try {
      localStorage.setItem(MOCK_STORAGE_KEY, JSON.stringify(prefs));
    } catch { /* ignore */ }
  }

  // ── List initialization ────────────────────────────────────────

  private async ensureList(): Promise<void> {
    if (this.listInitialized) return;

    try {
      const checkUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')`;
      const resp = await this.context.spHttpClient.get(
        checkUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );

      if (resp.status === 404) {
        await this.createList();
      } else {
        // List exists — make sure newer fields are present for previously-deployed lists.
        await this.ensureResultsViewField();
      }
      this.listInitialized = true;
    } catch {
      // If we can't check, attempt creation; if it already exists SP returns 400
      try {
        await this.createList();
      } catch { /* already exists */ }
      try {
        await this.ensureResultsViewField();
      } catch { /* already exists */ }
      this.listInitialized = true;
    }
  }

  private async ensureResultsViewField(): Promise<void> {
    const fieldUrl =
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')` +
      `/fields/getbytitle('ResultsView')`;
    const resp = await this.context.spHttpClient.get(
      fieldUrl,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata" } },
    );
    // SP returns 404 OR 400 when the field doesn't exist via getbytitle
    if (!resp.ok) {
      const fieldsUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/fields`;
      await this.context.spHttpClient.post(fieldsUrl, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
        },
        body: JSON.stringify({ FieldTypeKind: 2, Title: "ResultsView" }),
      });
    }
  }

  private async createList(): Promise<void> {
    const headers = {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
    };

    await this.context.spHttpClient.post(
      `${this.siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers,
        body: JSON.stringify({
          Title: this.listName,
          BaseTemplate: 100,
          Description:
            "SPECTRA user preferences — per-user favorites and recently viewed document IDs",
        }),
      },
    );

    const fieldsUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/fields`;

    // UserEmail column (Title is already built-in and will hold the email)
    // Add Favorites column
    await this.context.spHttpClient.post(fieldsUrl, SPHttpClient.configurations.v1, {
      headers,
      body: JSON.stringify({ FieldTypeKind: 3, Title: "Favorites" }),
    });

    // Add RecentlyViewed column
    await this.context.spHttpClient.post(fieldsUrl, SPHttpClient.configurations.v1, {
      headers,
      body: JSON.stringify({ FieldTypeKind: 3, Title: "RecentlyViewed" }),
    });

    // Add ResultsView column (single line of text: "table" | "tiles")
    await this.context.spHttpClient.post(fieldsUrl, SPHttpClient.configurations.v1, {
      headers,
      body: JSON.stringify({ FieldTypeKind: 2, Title: "ResultsView" }),
    });
  }

  // ── Public API ─────────────────────────────────────────────────

  async getUserPreferences(userEmail: string): Promise<IUserPreferences> {
    if (this.useMock) return this.mockLoad();

    const empty: IUserPreferences = {
      listItemId: null,
      favorites: [],
      recentlyViewed: [],
      resultsView: "table",
    };

    try {
      await this.ensureList();

      const encoded = userEmail.replace(/'/g, "''");
      const url =
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items` +
        `?$filter=Title eq '${encoded}'&$select=Id,Favorites,RecentlyViewed,ResultsView&$top=1`;

      let resp = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: "application/json;odata=nometadata" } },
      );

      // 400 means ResultsView column doesn't exist yet — retry without it
      if (resp.status === 400) {
        const fallbackUrl =
          `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items` +
          `?$filter=Title eq '${encoded}'&$select=Id,Favorites,RecentlyViewed&$top=1`;
        resp = await this.context.spHttpClient.get(
          fallbackUrl,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } },
        );
      }

      if (!resp.ok) return empty;

      const data = await resp.json();
      const items: IRawPrefsItem[] = data.value ?? [];

      if (items.length === 0) {
        // No SP row yet — merge in any locally-saved fallback (e.g. from a prior View-Only session)
        const local = loadLocalFallback(userEmail);
        if (!local.favorites?.length && !local.recentlyViewed?.length) return empty;
        return {
          ...empty,
          favorites: local.favorites ?? [],
          recentlyViewed: local.recentlyViewed ?? [],
          resultsView: local.resultsView ?? "table",
        };
      }

      const item = items[0];
      return {
        listItemId: item.Id,
        favorites: item.Favorites ? item.Favorites.split(";").filter(Boolean) : [],
        recentlyViewed: item.RecentlyViewed
          ? item.RecentlyViewed.split(";").filter(Boolean)
          : [],
        resultsView: item.ResultsView === "tiles" ? "tiles" : "table",
      };
    } catch {
      return empty;
    }
  }

  private async persistPreferences(
    userEmail: string,
    listItemId: number | null,
    favorites: string[],
    recentlyViewed: string[],
    resultsView: ResultsView,
  ): Promise<number | null> {
    const headers: Record<string, string> = {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
    };
    const body = JSON.stringify({
      Title: userEmail,
      Favorites: favorites.join(";"),
      RecentlyViewed: recentlyViewed.join(";"),
      ResultsView: resultsView,
    });

    if (listItemId === null) {
      const createUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;
      const resp = await this.context.spHttpClient.post(
        createUrl,
        SPHttpClient.configurations.v1,
        { headers, body },
      );
      if (resp.status === 403) return null; // caller will fall back to localStorage
      if (!resp.ok) return null;
      const data = await resp.json();
      return (data.Id as number) ?? null;
    }

    const updateUrl =
      `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${listItemId})`;
    const updateResp = await this.context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
      headers: { ...headers, "X-HTTP-Method": "MERGE", "IF-MATCH": "*" },
      body,
    });
    if (updateResp.status === 403) return null; // caller will fall back to localStorage
    return listItemId;
  }

  async toggleFavorite(
    userEmail: string,
    docId: string,
    current: IUserPreferences,
  ): Promise<{ prefs: IUserPreferences; added: boolean }> {
    const isFav = current.favorites.includes(docId);
    const newFavs = isFav
      ? current.favorites.filter((id) => id !== docId)
      : [...current.favorites, docId];

    if (this.useMock) {
      const prefs = { ...current, favorites: newFavs };
      this.mockSave(prefs);
      return { prefs, added: !isFav };
    }

    try {
      const newId = await this.persistPreferences(
        userEmail,
        current.listItemId,
        newFavs,
        current.recentlyViewed,
        current.resultsView,
      );
      const prefs: IUserPreferences = {
        ...current,
        listItemId: newId ?? current.listItemId,
        favorites: newFavs,
      };
      // If SP write was blocked (View Only), persist locally so the change survives refresh
      if (newId === null && current.listItemId === null) {
        saveLocalFallback(userEmail, prefs);
      }
      return { prefs, added: !isFav };
    } catch {
      return { prefs: current, added: false };
    }
  }

  async addRecentlyViewed(
    userEmail: string,
    docId: string,
    current: IUserPreferences,
  ): Promise<IUserPreferences> {
    const filtered = current.recentlyViewed.filter((id) => id !== docId);
    const newRV = [docId, ...filtered].slice(0, MAX_RECENTLY_VIEWED);

    if (this.useMock) {
      const prefs = { ...current, recentlyViewed: newRV };
      this.mockSave(prefs);
      return prefs;
    }

    try {
      const newId = await this.persistPreferences(
        userEmail,
        current.listItemId,
        current.favorites,
        newRV,
        current.resultsView,
      );
      const prefs = {
        ...current,
        listItemId: newId ?? current.listItemId,
        recentlyViewed: newRV,
      };
      if (newId === null && current.listItemId === null) {
        saveLocalFallback(userEmail, prefs);
      }
      return prefs;
    } catch {
      return current;
    }
  }

  async setResultsView(
    userEmail: string,
    value: ResultsView,
    current: IUserPreferences,
  ): Promise<IUserPreferences> {
    if (current.resultsView === value) return current;

    if (this.useMock) {
      const prefs = { ...current, resultsView: value };
      this.mockSave(prefs);
      return prefs;
    }

    try {
      const newId = await this.persistPreferences(
        userEmail,
        current.listItemId,
        current.favorites,
        current.recentlyViewed,
        value,
      );
      const prefs = {
        ...current,
        listItemId: newId ?? current.listItemId,
        resultsView: value,
      };
      if (newId === null && current.listItemId === null) {
        saveLocalFallback(userEmail, prefs);
      }
      return prefs;
    } catch {
      return current;
    }
  }
}
