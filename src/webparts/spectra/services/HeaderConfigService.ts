import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  FALLBACK_HEADER_CONFIG,
  IHeaderConfig,
  IHeaderLink,
} from "../interfaces/IHeaderConfig";
import {
  SP_LIST_NAMES,
  USE_MOCK as DEFAULT_USE_MOCK,
  HEADER_CONFIG_CACHE_KEY,
  HEADER_CONFIG_CACHE_TTL_MS,
} from "../config/config";

export class HeaderConfigService {
  private context: WebPartContext;
  private useMock: boolean;

  constructor(context: WebPartContext, useMock: boolean = DEFAULT_USE_MOCK) {
    this.context = context;
    this.useMock = useMock;
  }

  public async getConfig(): Promise<IHeaderConfig> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 250));
      return FALLBACK_HEADER_CONFIG;
    }

    const cached = this._readCache();
    try {
      const [settings, links] = await Promise.all([
        this._fetchSettings(),
        this._fetchLinks(),
      ]);

      const config: IHeaderConfig = {
        ...FALLBACK_HEADER_CONFIG,
        ...settings,
        links: links.length > 0 ? links : FALLBACK_HEADER_CONFIG.links,
      };

      this._writeCache(config);
      return config;
    } catch (error) {
      console.error("HeaderConfigService.getConfig:", error);
      if (cached) {
        return cached;
      }
      return FALLBACK_HEADER_CONFIG;
    }
  }

  private async _fetchSettings(): Promise<Partial<IHeaderConfig>> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const listName = SP_LIST_NAMES.headerSettings;
    const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$orderby=Modified desc&$top=50`;

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    if (!response.ok) {
      throw new Error(`Header settings list unavailable (${response.status})`);
    }

    const data = await response.json();
    const rows = (data.value as Record<string, unknown>[]).filter((item) =>
      this._toBoolean(this._read(item, ["IsEnabled"]), true),
    );

    const first = rows[0];
    if (!first) {
      return FALLBACK_HEADER_CONFIG;
    }

    return {
      parentAppName: this._toString(
        this._read(first, ["ParentAppName", "Title"]),
        FALLBACK_HEADER_CONFIG.parentAppName,
      ),
      parentAppUrl: this._toString(
        this._read(first, ["ParentAppUrl"]),
        FALLBACK_HEADER_CONFIG.parentAppUrl,
      ),
      parentLogoUrl: this._toString(
        this._read(first, ["ParentLogoUrl"]),
        FALLBACK_HEADER_CONFIG.parentLogoUrl || "",
      ),
      activeAppLabel: this._toString(
        this._read(first, ["ActiveAppLabel"]),
        FALLBACK_HEADER_CONFIG.activeAppLabel,
      ),
      activeAppUrl: this._toString(
        this._read(first, ["ActiveAppUrl"]),
        FALLBACK_HEADER_CONFIG.activeAppUrl || "#",
      ),
      activeAppIconName: this._toString(
        this._read(first, ["ActiveAppIconName"]),
        FALLBACK_HEADER_CONFIG.activeAppIconName || "fa-circle-dot",
      ),
    };
  }

  private async _fetchLinks(): Promise<IHeaderLink[]> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const listName = SP_LIST_NAMES.headerLinks;
    const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$top=200`;

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    if (!response.ok) {
      throw new Error(`Header links list unavailable (${response.status})`);
    }

    const data = await response.json();
    const rows = data.value as Record<string, unknown>[];

    return rows
      .filter((item) => this._toBoolean(this._read(item, ["IsActive"]), true))
      .map((item) => {
        const urlValue = this._toString(
          this._read(item, ["Url", "LinkUrl"]),
          "#",
        );

        const link: IHeaderLink = {
          id: String(this._read(item, ["Id"]) || this._toString(this._read(item, ["Title"]), "header-link")),
          label: this._toString(
            this._read(item, ["Label", "Title"]),
            "Link",
          ),
          url: urlValue,
          openInNewTab: this._toBoolean(
            this._read(item, ["OpenInNewTab"]),
            true,
          ),
          sortOrder: this._toNumber(
            this._read(item, ["SortOrder", "DisplayOrder"]),
            999,
          ),
          isActive: true,
          iconName: this._toString(
            this._read(item, ["IconName"]),
            "fa-arrow-up-right-from-square",
          ),
        };

        return link;
      })
      .sort((a, b) => a.sortOrder - b.sortOrder);
  }

  private _readCache(): IHeaderConfig | null {
    try {
      const raw = localStorage.getItem(HEADER_CONFIG_CACHE_KEY);
      if (!raw) return null;

      const parsed = JSON.parse(raw) as {
        data?: IHeaderConfig;
        expiry?: number;
      };

      if (!parsed.data || !parsed.expiry || Date.now() > parsed.expiry) {
        localStorage.removeItem(HEADER_CONFIG_CACHE_KEY);
        return null;
      }

      return parsed.data;
    } catch {
      return null;
    }
  }

  private _writeCache(data: IHeaderConfig): void {
    try {
      localStorage.setItem(
        HEADER_CONFIG_CACHE_KEY,
        JSON.stringify({
          data,
          expiry: Date.now() + HEADER_CONFIG_CACHE_TTL_MS,
        }),
      );
    } catch {
      // localStorage may be unavailable in private browsing modes.
    }
  }

  private _read(item: Record<string, unknown>, keys: string[]): unknown {
    for (const key of keys) {
      if (item[key] !== undefined && item[key] !== null) {
        return item[key];
      }
    }

    return undefined;
  }

  private _toString(value: unknown, fallback: string): string {
    if (typeof value === "string" && value.trim().length > 0) {
      return value.trim();
    }

    // SharePoint Hyperlink fields often come as objects: { Url, Description }
    if (value && typeof value === "object") {
      const record = value as Record<string, unknown>;
      const nestedUrl =
        this._asNonEmptyString(record.Url) ||
        this._asNonEmptyString(record.url) ||
        this._asNonEmptyString(record.Value) ||
        this._asNonEmptyString(record.value);

      if (nestedUrl) {
        return nestedUrl;
      }
    }

    return fallback;
  }

  private _asNonEmptyString(value: unknown): string | undefined {
    if (typeof value !== "string") {
      return undefined;
    }

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }

  private _toNumber(value: unknown, fallback: number): number {
    if (typeof value === "number" && !Number.isNaN(value)) {
      return value;
    }

    if (typeof value === "string") {
      const parsed = Number(value);
      if (!Number.isNaN(parsed)) {
        return parsed;
      }
    }

    return fallback;
  }

  private _toBoolean(value: unknown, fallback: boolean): boolean {
    if (typeof value === "boolean") {
      return value;
    }

    if (typeof value === "number") {
      return value !== 0;
    }

    if (typeof value === "string") {
      const normalized = value.trim().toLowerCase();
      if (normalized === "true" || normalized === "yes" || normalized === "1") {
        return true;
      }

      if (normalized === "false" || normalized === "no" || normalized === "0") {
        return false;
      }
    }

    return fallback;
  }
}
