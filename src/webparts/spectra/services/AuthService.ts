import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  MSGraphClientV3,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import { IAuthResponse } from "../interfaces/IAuthResponse";
import {
  USE_MOCK as DEFAULT_USE_MOCK,
  MOCK_ROLE,
  AD_GROUPS,
  SP_ROLE_GROUPS,
  USE_AD_GROUPS as DEFAULT_USE_AD_GROUPS,
  MOCK_GROUP_MEMBERSHIP,
  AUTH_DEBUG_LOGS,
  AUTH_LOG_PREFIX,
} from "../config/config";
import { getMockAuthResponse } from "../mocks/mockAuthResponse";

type RoleBucket = "admins" | "contributors" | "viewers";

interface ISpGroupUserEntry {
  Title?: string;
  PrincipalType?: number;
  Email?: string;
  LoginName?: string;
}

interface IRoleMap {
  admins: string[];
  contributors: string[];
  viewers: string[];
}

type AuthStageReporter = (
  stage: "authenticating" | "loadingUserInfo" | "assigningRole",
) => void;

export class AuthService {
  private context: WebPartContext;
  private useMock: boolean;
  private useAdGroups: boolean;
  private mockRole: "admin" | "contributor" | "viewer";
  private readonly maxRetryAttempts = 3;
  private readonly retryBaseDelayMs = 300;

  constructor(
    context: WebPartContext,
    useMock: boolean = DEFAULT_USE_MOCK,
    useAdGroups: boolean = DEFAULT_USE_AD_GROUPS,
    mockRole: "admin" | "contributor" | "viewer" = MOCK_ROLE,
  ) {
    this.context = context;
    this.useMock = useMock;
    this.useAdGroups = useAdGroups;
    this.mockRole = mockRole;
  }

  private _log(step: string, data?: unknown): void {
    if (!AUTH_DEBUG_LOGS) return;
    if (typeof data === "undefined") {
      console.info(`${AUTH_LOG_PREFIX} ${step}`);
      return;
    }
    console.info(`${AUTH_LOG_PREFIX} ${step}`, data);
  }

  private async _getGraphClient(): Promise<MSGraphClientV3> {
    return this.context.msGraphClientFactory.getClient("3");
  }

  private _normalizeGroupName(groupName: string): string {
    return groupName.trim().toLowerCase();
  }

  private _delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private _isRetryableHttpStatus(status: number): boolean {
    return status === 429 || status >= 500;
  }

  private async _withRetry<T>(
    operationName: string,
    operation: () => Promise<T>,
  ): Promise<T> {
    let lastError: unknown;

    for (let attempt = 1; attempt <= this.maxRetryAttempts; attempt += 1) {
      try {
        return await operation();
      } catch (error) {
        lastError = error;

        const errorStatus =
          error instanceof Error && "status" in error
            ? Number((error as Error & { status?: number }).status)
            : undefined;
        const isRetryableStatus =
          typeof errorStatus === "number" && this._isRetryableHttpStatus(errorStatus);

        if (attempt >= this.maxRetryAttempts || !isRetryableStatus) {
          break;
        }

        const delayMs = this.retryBaseDelayMs * attempt;
        this._log(`${operationName} retrying`, {
          attempt,
          nextAttempt: attempt + 1,
          delayMs,
          status: errorStatus,
        });
        await this._delay(delayMs);
      }
    }

    throw lastError instanceof Error ? lastError : new Error(`${operationName} failed`);
  }

  /**
   * Query a SharePoint role group's Users endpoint and return the AD
   * groups (PrincipalType = 4) listed in it. Used to dynamically build
   * the role-to-AD-group map without redeploying when admins add or
   * remove AD groups in SharePoint.
   */
  private async _getAdGroupsInSpRoleGroup(
    role: RoleBucket,
    spRoleGroupName: string,
  ): Promise<string[]> {
    return this._withRetry(`SharePoint role group members lookup [${role}]`, async () => {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const escapedGroupName = spRoleGroupName.replace(/'/g, "''");
      const url = `${siteUrl}/_api/web/sitegroups/getbyname('${encodeURIComponent(escapedGroupName)}')/Users?$select=Title,PrincipalType,Email,LoginName`;

      this._log("Loading SharePoint role group members", {
        role,
        spRoleGroupName,
        url,
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
          },
        },
      );

      this._log("SharePoint role group members response", {
        role,
        spRoleGroupName,
        status: response.status,
        statusText: response.statusText,
      });

      if (!response.ok) {
        if (this._isRetryableHttpStatus(response.status)) {
          throw new Error(`Retryable SharePoint status ${response.status}`);
        }

        this._log("SharePoint role group members lookup failed", {
          role,
          spRoleGroupName,
          status: response.status,
          statusText: response.statusText,
        });
        return [];
      }

      const data: { value?: ISpGroupUserEntry[] } = await response.json();
      const entries = data.value || [];
      const adGroupTitles = entries
        .filter((entry) => entry.PrincipalType === 4)
        .map((entry) => entry.Title)
        .filter((title): title is string => typeof title === "string" && title.length > 0);

      this._log("SharePoint role group members extracted", {
        role,
        spRoleGroupName,
        totalEntries: entries.length,
        adGroupCount: adGroupTitles.length,
        adGroupTitles,
        principalTypeBreakdown: entries.reduce<Record<string, number>>(
          (acc, entry) => {
            const key = String(entry.PrincipalType ?? "unknown");
            acc[key] = (acc[key] || 0) + 1;
            return acc;
          },
          {},
        ),
      });

      return adGroupTitles;
    });
  }

  /**
   * Build the role-to-AD-group map dynamically from SharePoint role
   * groups. Falls back to the hardcoded AD_GROUPS for any bucket that
   * fails to resolve or comes back empty.
   */
  private async _resolveDynamicRoleMap(): Promise<IRoleMap> {
    this._log("Resolving dynamic role map from SharePoint role groups", {
      spRoleGroups: SP_ROLE_GROUPS,
    });

    const buckets: RoleBucket[] = ["admins", "contributors", "viewers"];
    const resolved: IRoleMap = { admins: [], contributors: [], viewers: [] };

    for (const bucket of buckets) {
      try {
        const adGroups = await this._getAdGroupsInSpRoleGroup(
          bucket,
          SP_ROLE_GROUPS[bucket],
        );

        if (adGroups.length > 0) {
          resolved[bucket] = adGroups;
        } else {
          const fallback = [...AD_GROUPS[bucket]];
          this._log(
            "Dynamic role bucket empty; using hardcoded fallback",
            { bucket, fallbackAdGroups: fallback },
          );
          resolved[bucket] = fallback;
        }
      } catch (error) {
        const errorDetails =
          error instanceof Error
            ? { message: error.message, stack: error.stack }
            : { error };
        const fallback = [...AD_GROUPS[bucket]];
        this._log("Dynamic role bucket resolution threw; using fallback", {
          bucket,
          fallbackAdGroups: fallback,
          error: errorDetails,
        });
        resolved[bucket] = fallback;
      }
    }

    this._log("Dynamic role map resolved", {
      adminAdGroups: resolved.admins,
      contributorAdGroups: resolved.contributors,
      viewerAdGroups: resolved.viewers,
    });

    return resolved;
  }

  private _getNextGraphPath(nextLink?: string): string | undefined {
    if (!nextLink) {
      return undefined;
    }

    try {
      const url = new URL(nextLink);
      return `${url.pathname}${url.search}`.replace(/^\/v1\.0/i, "");
    } catch {
      return nextLink.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, "");
    }
  }

  private async _getCurrentUserTransitiveGroups(): Promise<
    Array<{ id?: string; displayName?: string }>
  > {
    return this._withRetry("Graph transitive group lookup", async () => {
      const client = await this._getGraphClient();
      const currentUserEmail = this.context.pageContext.user.email;
      const collectedGroups: Array<{ id?: string; displayName?: string }> = [];
      let requestPath:
        | string
        | undefined = "/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=999";
      let pageNumber = 0;

      this._log("Fetching current user transitive AD groups from Graph", {
        userEmail: currentUserEmail,
        requestPath,
      });

      while (requestPath) {
        pageNumber += 1;
        this._log("Graph transitive group request started", {
          userEmail: currentUserEmail,
          pageNumber,
          requestPath,
        });

        const response: {
          value?: Array<{ id?: string; displayName?: string }>;
          "@odata.nextLink"?: string;
        } = await client.api(requestPath).version("v1.0").get();

        const pageGroups = response.value || [];
        collectedGroups.push(...pageGroups);

        this._log("Graph transitive group page resolved", {
          userEmail: currentUserEmail,
          pageNumber,
          pageGroupCount: pageGroups.length,
          totalGroupCount: collectedGroups.length,
          sampleGroups: pageGroups.slice(0, 10).map((group) => ({
            id: group.id,
            displayName: group.displayName,
          })),
        });

        requestPath = this._getNextGraphPath(response["@odata.nextLink"]);
      }

      this._log("Graph transitive group fetch completed", {
        userEmail: currentUserEmail,
        totalGroupCount: collectedGroups.length,
        sampleGroups: collectedGroups.slice(0, 10).map((group) => ({
          id: group.id,
          displayName: group.displayName,
        })),
      });

      return collectedGroups;
    });
  }

  private async _checkAllGroupsForRole(
    role: IAuthResponse["roles"][number],
    groupNames: readonly string[],
    normalizedUserGroups: ReadonlySet<string>,
  ): Promise<boolean> {
    let hasMatch = false;

    this._log(`Evaluating role bucket [${role}]`, {
      groupsToCheck: groupNames,
      totalGroups: groupNames.length,
    });

    for (let idx = 0; idx < groupNames.length; idx += 1) {
      const groupName = groupNames[idx];
      this._log(`Starting group check [${role}]`, {
        groupName,
        groupIndex: idx + 1,
        totalGroups: groupNames.length,
      });

      const isMember = normalizedUserGroups.has(
        this._normalizeGroupName(groupName),
      );
      this._log(
        `Group check [${role}] ${groupName}: ${isMember ? "MEMBER" : "NOT MEMBER"}`,
      );
      if (isMember) {
        hasMatch = true;
      }
    }

    this._log(
      `Role bucket [${role}] result: ${hasMatch ? "MATCHED" : "NOT MATCHED"}`,
    );
    return hasMatch;
  }

  /**
   * Get the user's auth context — role, assets, exclusions.
   *
   * Mock mode: returns mock response based on MOCK_GROUP_MEMBERSHIP or MOCK_ROLE.
  * Live mode: checks Microsoft Graph for current-user transitive AD group
  * membership.
   */
  public async getUserContext(
    userEmail: string,
    reportStage?: AuthStageReporter,
  ): Promise<IAuthResponse> {
    const normalizedEmail = userEmail.trim().toLowerCase();
    reportStage?.("authenticating");
    this._log("Auth evaluation started", {
      userEmail: normalizedEmail,
      useMock: this.useMock,
      useAdGroups: this.useAdGroups,
    });

    // ── MOCK MODE / AD DISABLED MODE ─────────────────────────
    if (this.useMock || !this.useAdGroups) {
      await new Promise((resolve) => setTimeout(resolve, 300));

      const mappedGroups = MOCK_GROUP_MEMBERSHIP[normalizedEmail] || [];
      this._log("Using mock/AD-disabled auth path", {
        mappedGroupBuckets: mappedGroups,
        mockRoleFallback: this.mockRole,
      });

      if (mappedGroups.length > 0) {
        const roles: IAuthResponse["roles"] = [];

        if (mappedGroups.includes("admins")) roles.push("admin");
        if (mappedGroups.includes("contributors")) roles.push("contributor");
        if (mappedGroups.includes("viewers")) roles.push("viewer");

        if (roles.length > 0) {
          reportStage?.("assigningRole");
          this._log("Resolved roles from mock group buckets", { roles });
          return {
            roles,
            assets: [], // Asset-level filtering is planned for a future release
            exclusions: [], // Firewall exclusions are planned for a future release
            authVerified: true,
          };
        }
      }

      this._log("No mapped mock group buckets; using configured mock role", {
        mockRole: this.mockRole,
      });
      reportStage?.("assigningRole");
      return getMockAuthResponse(this.mockRole);
    }

    // ── LIVE MODE — Check Graph transitive AD group membership ─────────
    const roles: IAuthResponse["roles"] = [];

    try {
      const roleMap = await this._resolveDynamicRoleMap();

      this._log("Live auth role buckets configured", {
        adminGroups: roleMap.admins,
        contributorGroups: roleMap.contributors,
        viewerGroups: roleMap.viewers,
      });

      reportStage?.("loadingUserInfo");
      const userGroups = await this._getCurrentUserTransitiveGroups();
      const normalizedUserGroups = new Set(
        userGroups
          .map((group) => group.displayName)
          .filter((groupName): groupName is string => typeof groupName === "string")
          .map((groupName) => this._normalizeGroupName(groupName)),
      );

      this._log("Normalized Graph group names ready for evaluation", {
        normalizedGroupCount: normalizedUserGroups.size,
        sampleNormalizedGroups: Array.from(normalizedUserGroups).slice(0, 10),
      });

      reportStage?.("assigningRole");
      const isAdmin = await this._checkAllGroupsForRole(
        "admin",
        roleMap.admins,
        normalizedUserGroups,
      );
      if (isAdmin) roles.push("admin");

      const isContributor = await this._checkAllGroupsForRole(
        "contributor",
        roleMap.contributors,
        normalizedUserGroups,
      );
      if (isContributor) roles.push("contributor");

      const isViewer = await this._checkAllGroupsForRole(
        "viewer",
        roleMap.viewers,
        normalizedUserGroups,
      );
      if (isViewer) roles.push("viewer");
    } catch (error) {
      const errorDetails =
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error };
      this._log("Graph auth evaluation failed", errorDetails);
      console.error("AuthService: Failed to check group membership", error);
      throw error;
    }

    // Default to viewer if no group matched
    if (roles.length === 0) {
      roles.push("viewer");
      this._log("No AD group matches found; applying viewer fallback role");
    }

    this._log("Resolved raw roles before precedence", { roles });

    return {
      roles,
      assets: [], // All assets accessible — filtering by asset is Release 2
      exclusions: [], // Firewall exclusions — Release 2
      authVerified: true,
    };
  }

}
