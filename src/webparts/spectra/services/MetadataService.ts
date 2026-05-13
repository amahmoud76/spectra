import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  IMetadataOptions,
  IMetadataOption,
  IDiseaseAreaStrategyRelationship,
  IProjectPaidRelationship,
  EMPTY_METADATA,
} from "../interfaces/IMetadataOptions";
import {
  USE_MOCK as DEFAULT_USE_MOCK,
  SP_LIST_NAMES,
  METADATA_CACHE_KEY,
  METADATA_CACHE_TTL_MS,
} from "../config/config";
import { mockMetadataOptions } from "../mocks/mockMetadata";
import { getTime } from "date-fns";

export class MetadataService {
  private context: WebPartContext;
  private useMock: boolean;

  constructor(context: WebPartContext, useMock: boolean = DEFAULT_USE_MOCK) {
    this.context = context;
    this.useMock = useMock;
  }

  /**
   * Load all metadata options from reference lists.
   * Returns cached data if available and not expired.
   */
  public async getOptions(): Promise<IMetadataOptions> {
    // ── MOCK MODE ─────────────────────────────────────────────
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 300));
      return mockMetadataOptions;
    }

    // ── Check cache ───────────────────────────────────────────
    const cached = this._readCache();
    if (cached) {
      // Keep FilterPanel indication source strictly tied to SPECTRA_Indications.
      // If cache was populated during a transient list failure, self-heal by
      // re-fetching indications before returning cached metadata.
      if (cached.indications.length === 0) {
        try {
          const refreshedIndications = await this._fetchSimpleList(
            SP_LIST_NAMES.indications,
            true,
          );

          if (refreshedIndications.length > 0) {
            const repairedCache: IMetadataOptions = {
              ...cached,
              indications: refreshedIndications,
            };
            this._writeCache(repairedCache);
            return repairedCache;
          }
        } catch {
          // Keep existing cached metadata when indication refresh fails.
        }
      }

      return cached;
    }

    // ── LIVE MODE — Fetch all lists in parallel ───────────────
    try {
      const [
        assets,
        documentTypes,
        diseaseAreaData,
        therapeuticAreas,
        subTherapeuticAreas,
        indications,
        lineOfTherapy,
        projectPaidRelationships,
      ] = await Promise.all([
        this._withMetadataFallback("assets", () => this._fetchAssets(), []),
        this._withMetadataFallback(
          "documentTypes",
          () => this._fetchSimpleList(SP_LIST_NAMES.documentTypes, true),
          [],
        ),
        this._withMetadataFallback(
          "diseaseAreaStrategies",
          () => this._fetchDiseaseAreaStrategies(),
          { options: [], relationships: [] },
        ),
        this._withMetadataFallback(
          "therapeuticAreas",
          () => this._fetchTherapeuticAreas(),
          [],
        ),
        this._withMetadataFallback(
          "subTherapeuticAreas",
          () => this._fetchSubTherapeuticAreas(),
          [],
        ),
        this._withMetadataFallback(
          "indications",
          () => this._fetchSimpleList(SP_LIST_NAMES.indications, true),
          [],
        ),
        this._withMetadataFallback(
          "lineOfTherapy",
          () => this._fetchSimpleList(SP_LIST_NAMES.lineOfTherapy, false),
          [],
        ),
        this._withMetadataFallback(
          "projectPaidRelationships",
          () => this._fetchProjectPaidRelationships(),
          [],
        ),
      ]);

      const options: IMetadataOptions = {
        assets,
        documentTypes,
        diseaseAreaStrategies: diseaseAreaData.options,
        diseaseAreaStrategyRelationships: diseaseAreaData.relationships,
        therapeuticAreas,
        subTherapeuticAreas,
        indications,
        lineOfTherapy,
        projectPaidRelationships,
      };

      this._writeCache(options);
      return options;
    } catch (error) {
      console.error("MetadataService.getOptions:", error);
      return EMPTY_METADATA;
    }
  }

  // ─────────────────────────────────────────────────────────────
  // PRIVATE — Fetch individual lists
  // ─────────────────────────────────────────────────────────────

  /**
   * Fetch a simple list with Title and optional search-tokens column.
   * searchTokensColumn — actual column name in the list (varies per list).
   */
  private async _fetchSimpleList(
    listName: string,
    includeSearchTokens: boolean,
    searchTokensColumn: string = "SEARCH_TOKENS",
  ): Promise<IMetadataOption[]> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const tokenColumnCandidates: string[] = includeSearchTokens
      ? Array.from(
          new Set([
            searchTokensColumn,
            "SEARCH_TOKENS",
          ]),
        )
      : [];

    const attempts: Array<string | null> = includeSearchTokens
      ? [...tokenColumnCandidates, null]
      : [null];

    let lastStatus: number | null = null;

    for (const tokenColumn of attempts) {
      const select = tokenColumn ? `Title,${tokenColumn}` : "Title";
      const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${select}&$orderby=Title&$top=5000`;

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
        );

      if (response.ok) {
        const data = await response.json();
        return (data.value as Array<Record<string, string>>).map((item) => ({
          value: item.Title,
          searchTokens:
            tokenColumn && item[tokenColumn]
              ? item[tokenColumn]
                  .split(",")
                  .map((t: string) => t.trim())
                  .filter(Boolean)
              : [],
        }));
      }

      lastStatus = response.status;

      // 400 usually means a bad field name in $select; try next fallback.
      if (response.status === 400 && includeSearchTokens) {
        continue;
      }

      throw new Error(
        `MetadataService: failed to load "${listName}" (${response.status})`,
      );
    }

    throw new Error(
      `MetadataService: failed to load "${listName}" (${lastStatus ?? "unknown"})`,
    );
  }

  /**
   * Fetch SPECTRA_Assets — columns: Title, SEARCH_TOKENS, ASSET_NUMBER.
   * Uses Title as the display value and ASSET_NUMBER as a search token.
   */
  private async _fetchAssets(): Promise<IMetadataOption[]> {
    const responses = await this._fetchListItemsWithFallback(
      SP_LIST_NAMES.assets,
      [
        {
          select: ["Title", "SEARCH_TOKENS", "ASSET_NUMBER"],
          orderBy: "Title",
        },
        { select: ["Title", "ASSET_NUMBER"], orderBy: "Title" },
      ],
    );

    const seen = new Set<string>();

    return responses
      .map((item) => {
        const value = this._readString(item, ["Title", "ASSET_NUMBER"]);
        const rawSearchTokens = this._readString(item, ["SEARCH_TOKENS"]);
        const assetNumber = this._readString(item, ["ASSET_NUMBER"]);
        const searchTokens = Array.from(
          new Set(
            [
              ...rawSearchTokens
                .split(/\s*[|,]\s*/)
                .map((token) => token.trim())
                .filter(Boolean),
              assetNumber,
            ].filter(Boolean),
          ),
        );

        return {
          value,
          searchTokens,
        };
      })
      .filter((item) => {
        if (!item.value || seen.has(item.value)) {
          return false;
        }

        seen.add(item.value);
        return true;
      });
  }

  /**
   * Fetch SPECTRA_TherapeuticAreas — columns: Title, ABBREVIATION.
   * Builds searchTokens from both Title and Abbreviation for synonym search.
   * Example: searching "AES" finds "Aesthetics" and vice versa.
   */
  private async _fetchTherapeuticAreas(): Promise<IMetadataOption[]> {
    const responses = await this._fetchListItemsWithFallback(
      SP_LIST_NAMES.therapeuticAreas,
      [
        { select: ["Title", "ABBREVIATION"], orderBy: "Title" },
        { select: ["Title"], orderBy: "Title" },
      ],
    );

    return responses.map((item) => {
      const title = this._readString(item, ["Title"]);
      const abbreviation = this._readString(item, ["ABBREVIATION"]);
      const searchTokens: string[] = [];

      // Always include abbreviation if present
      if (abbreviation) {
        searchTokens.push(abbreviation);
      }

      return {
        value: title,
        searchTokens,
      };
    });
  }

  /**
   * Fetch SPECTRA_SubTherapeuticAreas — columns: Title, THERAPEUTIC_AREA, Abbreviation.
   * Uses Title as the display value.
   */
  private async _fetchSubTherapeuticAreas(): Promise<IMetadataOption[]> {
    const responses = await this._fetchListItemsWithFallback(
      SP_LIST_NAMES.subTherapeuticAreas,
      [
        { select: ["Title", "THERAPEUTIC_AREA"], orderBy: "Title" },
        { select: ["Title"], orderBy: "Title" },
      ],
    );

    return responses.map((item) => ({
      value: this._readString(item, ["Title"]),
      searchTokens: this._readString(item, ["THERAPEUTIC_AREA"])
        ? [this._readString(item, ["THERAPEUTIC_AREA"])]
        : [],
    }));
  }

  /**
   * Fetch SPECTRA_DAS — columns: Title, THERAPEUTIC_AREA.
   * Returns both a unique DAS option list and the raw TA -> DAS relationship rows.
   */
  private async _fetchDiseaseAreaStrategies(): Promise<{
    options: IMetadataOption[];
    relationships: IDiseaseAreaStrategyRelationship[];
  }> {
    const responses = await this._fetchListItemsWithFallback(
      SP_LIST_NAMES.diseaseAreaStrategies,
      [
        { select: ["Title", "THERAPEUTIC_AREA"], orderBy: "Title" },
        { select: ["Title"], orderBy: "Title" },
      ],
    );

    const relationships = responses
      .map((item) => ({
        value: this._readString(item, ["Title"]),
        therapeuticArea: this._readString(item, ["THERAPEUTIC_AREA"]),
      }))
      .filter((item) => item.value)
      .map((item) => ({
        value: item.value,
        therapeuticArea: item.therapeuticArea,
      }));

    const optionMap = new Map<string, IMetadataOption>();
    relationships.forEach((item) => {
      const existing = optionMap.get(item.value);
      if (existing) {
        if (
          item.therapeuticArea &&
          !existing.searchTokens.includes(item.therapeuticArea)
        ) {
          existing.searchTokens.push(item.therapeuticArea);
        }
        return;
      }

      optionMap.set(item.value, {
        value: item.value,
        searchTokens: item.therapeuticArea ? [item.therapeuticArea] : [],
      });
    });

    return {
      options: Array.from(optionMap.values()),
      relationships,
    };
  }

  /**
   * Fetch SPECTRA_ProjectPAID — cascading relationship data.
   * Returns all rows — used for both PAID dropdown values and cascading filters.
   */
  private async _fetchProjectPaidRelationships(): Promise<
    IProjectPaidRelationship[]
  > {
    const responses = await this._fetchListItemsWithFallback(
      SP_LIST_NAMES.projectPaid,
      [
        {
          select: [
            "Title",
            "THERAPEUTIC_AREA",
            "ASSET_NUMBER",
            "INDICATION",
            "SUB_TA",
            "LINE_OF_THERAPY",
          ],
          orderBy: "Title",
        },
        {
          select: [
            "PROJECT_PAID",
            "THERAPEUTIC_AREA",
            "ASSET_NUMBER",
            "INDICATION",
            "SUB_TA",
            "LINE_OF_THERAPY",
          ],
          orderBy: "PROJECT_PAID",
        },
        {
          select: [
            "Title",
            "THERAPEUTIC_AREA",
            "ASSET_NUMBER",
            "INDICATION",
          ],
          orderBy: "Title",
        },
        { select: ["Title"], orderBy: "Title" },
      ],
    );

    return responses
      .map((item) => ({
        projectPaid: this._readString(item, ["Title", "PROJECT_PAID"]),
        therapeuticArea: this._readString(item, ["THERAPEUTIC_AREA"]),
        assetNumber: this._readString(item, ["ASSET_NUMBER"]),
        indication: this._readString(item, ["INDICATION"]),
        subTherapeuticArea: this._readString(item, ["SUB_TA"]),
        lineOfTherapy: this._readString(item, ["LINE_OF_THERAPY"]),
      }))
      .filter((item) => item.projectPaid);
  }

  private async _withMetadataFallback<T>(
    label: string,
    load: () => Promise<T>,
    fallback: T,
  ): Promise<T> {
    try {
      return await load();
    } catch (error) {
      console.error(`MetadataService: ${label} fallback`, error);
      return fallback;
    }
  }

  private async _fetchListItemsWithFallback(
    listName: string,
    attempts: Array<{ select: string[]; orderBy: string }>,
  ): Promise<Array<Record<string, unknown>>> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    let lastStatus: number | null = null;

    for (const attempt of attempts) {
      const select = Array.from(new Set(attempt.select)).join(",");
      const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${select}&$orderby=${attempt.orderBy}&$top=5000`;

      const response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
      );

      if (response.ok) {
        const data = await response.json();
        return data.value as Array<Record<string, unknown>>;
      }

      lastStatus = response.status;
      if (response.status !== 400) {
        throw new Error(
          `MetadataService: failed to load "${listName}" (${response.status})`,
        );
      }
    }

    throw new Error(
      `MetadataService: failed to load "${listName}" (${lastStatus ?? "unknown"})`,
    );
  }

  private _readString(item: Record<string, unknown>, keys: string[]): string {
    for (const key of keys) {
      const value = item[key];
      if (typeof value === "string" && value.trim().length > 0) {
        return value.trim();
      }
    }

    return "";
  }

  // ─────────────────────────────────────────────────────────────
  // CACHE
  // ─────────────────────────────────────────────────────────────

  private _readCache(): IMetadataOptions | null {
    try {
      const raw = localStorage.getItem(METADATA_CACHE_KEY);
      if (!raw) return null;
      const { data, expiry } = JSON.parse(raw);
      if (getTime(new Date()) > expiry) {
        localStorage.removeItem(METADATA_CACHE_KEY);
        return null;
      }
      return data as IMetadataOptions;
    } catch {
      return null;
    }
  }

  private _writeCache(options: IMetadataOptions): void {
    try {
      localStorage.setItem(
        METADATA_CACHE_KEY,
        JSON.stringify({
          data: options,
          expiry: getTime(new Date()) + METADATA_CACHE_TTL_MS,
        }),
      );
    } catch {
      // localStorage quota exceeded — proceed without caching
    }
  }
}
