import { useState, useEffect, useCallback, useRef } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  IProcessedAuth,
  ApiRole,
  EffectiveRole,
  AuthStartupStage,
} from "../interfaces/IAuthResponse";
import { AuthService } from "../services/AuthService";
import { captureAndLogError } from "../services/errorLogService";
import { getCachedAuth, setCachedAuth } from "../utils/cacheHelper";
import { AUTH_DEBUG_LOGS, AUTH_LOG_PREFIX } from "../config/config";

interface IAuthFetchState {
  effectiveRole: EffectiveRole;
  assets: string[];
  exclusions: string[];
  isLoading: boolean;
  isError: boolean;
}

const DEV_ROLE_OVERRIDE_KEY = "spectra.dev.overrideRole";

const parseStoredRole = (value?: string): EffectiveRole | undefined => {
  if (value === "admin" || value === "contributor" || value === "viewer") {
    return value;
  }
  return undefined;
};

const getNextRole = (role: EffectiveRole): EffectiveRole => {
  if (role === "admin") return "contributor";
  if (role === "contributor") return "viewer";
  return "admin";
};

/**
 * Priority: admin > contributor > viewer
 * If user is in multiple groups, highest priority wins.
 */
const deriveEffectiveRole = (roles: ApiRole[]): EffectiveRole => {
  if (roles.includes("admin")) return "admin" as EffectiveRole;
  if (roles.includes("contributor")) return "contributor" as EffectiveRole;
  return "viewer" as EffectiveRole;
};

export const useAuth = (
  context: WebPartContext,
  userEmail: string,
  enableDevRoleSwitch: boolean,
  useMock: boolean,
  useAdGroups: boolean,
  mockRole: EffectiveRole,
): IProcessedAuth => {
  const logAuth = useCallback((step: string, data?: unknown): void => {
    if (!AUTH_DEBUG_LOGS) return;
    if (typeof data === "undefined") {
      console.info(`${AUTH_LOG_PREFIX} ${step}`);
      return;
    }
    console.info(`${AUTH_LOG_PREFIX} ${step}`, data);
  }, []);

  const [state, setState] = useState<IAuthFetchState>({
    effectiveRole: "viewer" as EffectiveRole,
    assets: [] as string[],
    exclusions: [] as string[],
    isLoading: true,
    isError: false,
  });

  const [devOverrideRole, setDevOverrideRole] = useState<
    EffectiveRole | undefined
  >(undefined);
  const [startupStage, setStartupStage] = useState<AuthStartupStage>("idle");
  const [authRefreshToken, setAuthRefreshToken] = useState(0);
  const autoRetryCountRef = useRef(0);
  const retryTimerRef = useRef<number | null>(null);

  const MAX_AUTO_RETRIES = 3;
  const RETRY_DELAY_MS = 7500;

  const retryAuth = useCallback(() => {
    autoRetryCountRef.current = 0;
    setAuthRefreshToken((current) => current + 1);
  }, []);

  const logAuthFailure = useCallback(
    (message: string, error?: unknown): void => {
      void captureAndLogError(error instanceof Error ? error : new Error(message), {
        component: "useAuth",
        errorType: "Auth Verification Failure",
        featureFunction: "useAuth",
        userEmail,
        userRole: "viewer",
        userAction: "Auth lookup",
        additionalContext: {
          message,
          userEmail: userEmail.trim().toLowerCase(),
          useMock,
          useAdGroups,
          mockRole,
          authRefreshToken,
        },
      }).catch(() => undefined);
    },
    [authRefreshToken, mockRole, useAdGroups, useMock, userEmail],
  );

  useEffect(() => {
    if (!enableDevRoleSwitch) {
      setDevOverrideRole(undefined);
      return;
    }

    try {
      const stored = parseStoredRole(
        localStorage.getItem(DEV_ROLE_OVERRIDE_KEY) || undefined,
      );
      setDevOverrideRole(stored);
    } catch {
      setDevOverrideRole(undefined);
    }
  }, [enableDevRoleSwitch]);

  useEffect(() => {
    if (!userEmail) return;

    const fetchAuth = async (): Promise<void> => {
      setState((previous) => ({ ...previous, isLoading: true, isError: false }));
      setStartupStage("authenticating");
      logAuth("useAuth fetch started", {
        userEmail: userEmail.trim().toLowerCase(),
        useMock,
        useAdGroups,
        mockRole,
      });

      // Check cache first
      const cached = authRefreshToken === 0 ? getCachedAuth(userEmail) : null;
      if (cached) {
        const derivedRole = deriveEffectiveRole(cached.roles);
        setStartupStage("ready");
        logAuth("Cache hit for user auth", {
          roles: cached.roles,
          derivedRole,
        });

        setState({
          effectiveRole: derivedRole,
          assets: cached.assets,
          exclusions: cached.exclusions,
          isLoading: false,
          isError: false,
        });
        return;
      }

      logAuth("Cache miss for user auth; fetching live auth context");

      try {
        const service = new AuthService(
          context,
          useMock,
          useAdGroups,
          mockRole,
        );
        const response = await service.getUserContext(userEmail, (stage) => {
          setStartupStage(stage);
        });
        const derivedRole = deriveEffectiveRole(response.roles);
        setStartupStage("ready");

        if (!response.authVerified) {
          logAuth("Auth verification failed; refusing to cache fallback result", {
            roles: response.roles,
            derivedRole,
          });
          logAuthFailure("Auth verification failed; returned fallback viewer state", {
            roles: response.roles,
            derivedRole,
          });
          setState({
            effectiveRole: "viewer" as EffectiveRole,
            assets: response.assets,
            exclusions: response.exclusions,
            isLoading: false,
            isError: true,
          });
          return;
        }

        autoRetryCountRef.current = 0;
        setCachedAuth(userEmail, response);

        logAuth("Live auth response received", {
          roles: response.roles,
          derivedRole,
        });

        setState({
          effectiveRole: derivedRole,
          assets: response.assets,
          exclusions: response.exclusions,
          isLoading: false,
          isError: false,
        });
      } catch (error) {
        if (AUTH_DEBUG_LOGS) {
          console.error("useAuth: Failed to fetch user context", error);
        }

        autoRetryCountRef.current += 1;

        if (autoRetryCountRef.current <= MAX_AUTO_RETRIES) {
          logAuth("Auth fetch failed; scheduling auto-retry", {
            attempt: autoRetryCountRef.current,
            maxRetries: MAX_AUTO_RETRIES,
          });
          setStartupStage("retrying");
          retryTimerRef.current = window.setTimeout(() => {
            retryTimerRef.current = null;
            setAuthRefreshToken((n) => n + 1);
          }, RETRY_DELAY_MS);
          return;
        }

        autoRetryCountRef.current = 0;
        logAuth("Auth fetch failed after all retries; falling back to viewer", { error });
        logAuthFailure("Auth fetch failed after retries", error);
        setStartupStage("ready");
        setState((prev) => ({
          ...prev,
          effectiveRole: "viewer" as EffectiveRole,
          isLoading: false,
          isError: true,
        }));
      }
    };

    void fetchAuth();

    return () => {
      if (retryTimerRef.current !== null) {
        window.clearTimeout(retryTimerRef.current);
        retryTimerRef.current = null;
      }
    };
  }, [context, userEmail, useMock, useAdGroups, mockRole, logAuth, authRefreshToken]);

  useEffect(() => {
    if (state.isLoading) return;

    const finalRole =
      enableDevRoleSwitch && devOverrideRole
        ? devOverrideRole
        : state.effectiveRole;

    logAuth("Final role resolved", {
      derivedRole: state.effectiveRole,
      devOverrideRole: devOverrideRole ?? null,
      finalRole,
    });
  }, [state.isLoading, state.effectiveRole, devOverrideRole, enableDevRoleSwitch, logAuth]);

  const cycleDevRole = useCallback(() => {
    if (!enableDevRoleSwitch) return;

    setDevOverrideRole((previousRole) => {
      const baseRole = previousRole ?? state.effectiveRole;
      const nextRole = getNextRole(baseRole);

      try {
        localStorage.setItem(DEV_ROLE_OVERRIDE_KEY, nextRole);
      } catch {
        // Ignore storage failures in constrained browser modes.
      }

      if (AUTH_DEBUG_LOGS) {
        console.info(`SPECTRA dev role override active: ${nextRole}`);
      }
      return nextRole;
    });
  }, [enableDevRoleSwitch, state.effectiveRole]);

  const clearDevRoleOverride = useCallback(() => {
    if (!enableDevRoleSwitch) return;

    try {
      localStorage.removeItem(DEV_ROLE_OVERRIDE_KEY);
    } catch {
      // Ignore storage failures in constrained browser modes.
    }

    setDevOverrideRole(undefined);
  }, [enableDevRoleSwitch]);

  return {
    ...state,
    startupStage,
    effectiveRole:
      enableDevRoleSwitch && devOverrideRole
        ? devOverrideRole
        : state.effectiveRole,
    devOverrideRole,
    isDevRoleSwitchEnabled: enableDevRoleSwitch,
    cycleDevRole,
    clearDevRoleOverride,
    retryAuth,
  };
};
