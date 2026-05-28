// ─────────────────────────────────────────────────────────────────
// Auth response — role determined by AD Group membership
//
// AD Group mapping:
//   APP-INFORM-PipelineAnalyticsOps, APP-QlikSense-PPM-BTS → 'admin'
//   APP-IlluminationHub-Contributor                         → 'contributor'
//   APP-IlluminationHub-AdHoc, APP-QlikSense-PMO-Roster-PROD,
//   APP-IlluminationHub-LT                                  → 'viewer'
//   (default if no group matched)
// ─────────────────────────────────────────────────────────────────

// Role as returned by AuthService
export type ApiRole = "admin" | "contributor" | "viewer";

// Effective role after priority resolution
// (if user is in multiple groups, highest priority wins)
export type EffectiveRole = "admin" | "contributor" | "viewer";

export type AuthStartupStage =
  | "idle"
  | "authenticating"
  | "loadingUserInfo"
  | "assigningRole"
  | "retrying"
  | "ready";

// Raw response from AuthService
export interface IAuthResponse {
  roles: ApiRole[]; // All roles the user belongs to
  assets: string[]; // Assets the user has access to
  exclusions: string[]; // Assets excluded (firewall — Release 2)
  authVerified: boolean; // True when the live auth lookup completed successfully
}

// Processed auth state used by useAuth hook and all components
export interface IProcessedAuth {
  effectiveRole: EffectiveRole;
  assets: string[];
  exclusions: string[];
  isLoading: boolean;
  isError: boolean;
  startupStage: AuthStartupStage;
  devOverrideRole?: EffectiveRole;
  isDevRoleSwitchEnabled: boolean;
  cycleDevRole: () => void;
  clearDevRoleOverride: () => void;
  retryAuth: () => void;
}
