export interface IErrorLog {
  Title: string;
  ErrorMessage: string;
  ErrorType: string;
  ComponentPage: string;
  ErrorTimestamp: string;
  UserEmail: string;
  UserDisplayName: string;
  SeverityLevel: string;
  StackTrace: string;
  BrowserInfo: string;
  UserAction?: string;
  AdditionalContext?: string;
  FeatureFunction?: string;
  UserRole?: string;
}

export interface IErrorLogResult {
  success: boolean;
  itemId?: number;
  message?: string;
}

const ERROR_LOG_LIST_NAME = "SPECTRA_ErrorLog";
const ERROR_LOG_LIST_SERVER_RELATIVE_URL = "/teams/SPECTRA/Lists/SPECTRA_ErrorLog";
const ERROR_LOG_LIST_ABSOLUTE_URL =
  "https://abbvie.sharepoint.com/teams/SPECTRA/Lists/SPECTRA_ErrorLog/AllItems.aspx";
const SPECTRA_WEB_SERVER_RELATIVE_URL =
  ERROR_LOG_LIST_SERVER_RELATIVE_URL.split("/Lists/")[0] || "/teams/SPECTRA";
const ERROR_LOG_LIST_NAME_CANDIDATES = [
  ERROR_LOG_LIST_NAME,
  "SPECTRA ErrorLog",
  "SPECTRA Error Log",
  "Error Log",
];

const getWebRootUrl = (): string => {
  return `${window.location.origin}${SPECTRA_WEB_SERVER_RELATIVE_URL}`;
};

const encodeODataString = (value: string): string => value.replace(/'/g, "''");

const getListByPathApi = (serverRelativePath: string): string => {
  const encodedPath = encodeURIComponent(serverRelativePath);
  return `getList('${encodedPath}')`;
};

const getListByTitleApi = (title: string): string => {
  const escapedTitle = encodeODataString(title);
  return `lists/getByTitle('${escapedTitle}')`;
};

const fetchJson = async (
  url: string,
  init?: RequestInit,
): Promise<unknown> => {
  const response = await fetch(url, {
    credentials: "same-origin",
    headers: {
      Accept: "application/json;odata=nometadata",
      ...(init?.headers || {}),
    },
    ...init,
  });

  const responseText = await response.text();

  if (!response.ok) {
    throw new Error(
      `HTTP ${response.status} calling ${url} :: ${responseText.trim().slice(0, 300)}`,
    );
  }

  if (response.status === 204 || !responseText.trim()) {
    return {};
  }

  const trimmed = responseText.trim();
  if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
    return JSON.parse(trimmed) as unknown;
  }

  // SharePoint can return Atom XML for successful POSTs. Return raw text instead of failing.
  return trimmed;
};

const extractValueArray = <T>(payload: unknown): T[] => {
  if (!payload || typeof payload !== "object") {
    return [];
  }

  const root = payload as {
    value?: T[];
    d?: { results?: T[] };
  };

  if (Array.isArray(root.value)) {
    return root.value;
  }

  const verboseResults = root.d?.results;
  if (Array.isArray(verboseResults)) {
    return verboseResults;
  }

  return [];
};

const extractItemId = (payload: unknown): number | undefined => {
  if (!payload) {
    return undefined;
  }

  if (typeof payload === "object") {
    const root = payload as {
      Id?: number;
      ID?: number;
      d?: { Id?: number; ID?: number };
    };

    return root.Id || root.ID || root.d?.Id || root.d?.ID;
  }

  if (typeof payload === "string") {
    const idMatch = payload.match(/<d:Id[^>]*>(\d+)<\/d:Id>/i);
    if (idMatch && idMatch[1]) {
      return Number(idMatch[1]);
    }
  }

  return undefined;
};

const getRequestDigest = async (webUrl: string): Promise<string> => {
  const result = (await fetchJson(`${webUrl}/_api/contextinfo`, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
    },
  })) as {
    FormDigestValue?: string;
    d?: { GetContextWebInformation?: { FormDigestValue?: string } };
  };

  const digest =
    result.FormDigestValue || result.d?.GetContextWebInformation?.FormDigestValue;

  if (!digest) {
    throw new Error("Unable to retrieve SharePoint request digest");
  }

  return digest;
};

const getListFields = async (
  webUrl: string,
  listApi: string,
): Promise<Set<string>> => {
  const fieldsResult = await fetchJson(
    `${webUrl}/_api/web/${listApi}/fields?$select=InternalName`,
  );
  const fields = extractValueArray<{ InternalName: string }>(fieldsResult);
  return new Set<string>(fields.map((field) => field.InternalName));
};

const isNotFoundError = (error: unknown): boolean => {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes("HTTP 404") || message.includes("[404]");
};

const resolveErrorLogList = async (): Promise<{ listTitle: string; listApi: string; fields: Set<string> }> => {
  const webUrl = getWebRootUrl();

  // Prefer the known server-relative list URL first.
  try {
    const listApi = getListByPathApi(ERROR_LOG_LIST_SERVER_RELATIVE_URL);
    const availableFields = await getListFields(webUrl, listApi);
    return { listTitle: ERROR_LOG_LIST_NAME, listApi, fields: availableFields };
  } catch (error) {
    if (!isNotFoundError(error)) {
      throw error;
    }
  }

  for (const candidate of ERROR_LOG_LIST_NAME_CANDIDATES) {
    try {
      const listApi = getListByTitleApi(candidate);
      const availableFields = await getListFields(webUrl, listApi);
      return { listTitle: candidate, listApi, fields: availableFields };
    } catch (error) {
      if (!isNotFoundError(error)) {
        throw error;
      }
    }
  }

  // Fallback discovery: try finding a likely list by title keywords.
  const allVisibleLists = await fetchJson(
    `${webUrl}/_api/web/lists?$select=Title,Hidden&$filter=Hidden eq false`,
  );

  const listTitles = extractValueArray<{ Title?: string }>(allVisibleLists)
    .map((list: { Title?: string }) => list.Title)
    .filter((title: string | undefined): title is string => !!title);

  const spectraErrorList = listTitles.find((title) => {
    const normalized = title.toLowerCase();
    return normalized.includes("spectra") && normalized.includes("error");
  });

  const genericErrorList = listTitles.find((title) =>
    title.toLowerCase().includes("error"),
  );

  const discoveredListTitle = spectraErrorList || genericErrorList;
  if (discoveredListTitle) {
    const listApi = getListByTitleApi(discoveredListTitle);
    const availableFields = await getListFields(webUrl, listApi);
    return { listTitle: discoveredListTitle, listApi, fields: availableFields };
  }

  const attemptedTitles = ERROR_LOG_LIST_NAME_CANDIDATES.join(", ");
  const siteUrl = window.location.origin + window.location.pathname;
  throw new Error(
    `Error log list not found. Tried list URL ${ERROR_LOG_LIST_SERVER_RELATIVE_URL}, exact titles [${attemptedTitles}], keyword discovery on site ${siteUrl}. Confirm list exists at ${ERROR_LOG_LIST_ABSOLUTE_URL}`,
  );
};

const firstNonEmpty = (values: Array<string | undefined>): string | undefined => {
  return values.find((value) => typeof value === "string" && value.trim().length > 0);
};

const setIfFieldExists = (
  payload: Record<string, string>,
  fields: Set<string>,
  candidates: string[],
  value: string | undefined,
): void => {
  if (!value) return;
  const target = candidates.find((candidate) => fields.has(candidate));
  if (target) {
    payload[target] = value;
  }
};

const buildListPayload = (
  errorLog: IErrorLog,
  fields: Set<string>,
): Record<string, string> => {
  const payload: Record<string, string> = {};

  const titleValue = firstNonEmpty([errorLog.Title]) || `Error-${Date.now()}`;
  payload.Title = titleValue;

  setIfFieldExists(payload, fields, ["ErrorMessage", "Error_x0020_Message"], errorLog.ErrorMessage);
  setIfFieldExists(payload, fields, ["ErrorType", "Error_x0020_Type"], errorLog.ErrorType);
  setIfFieldExists(payload, fields, ["ComponentPage", "Component_x002f_Page"], errorLog.ComponentPage);
  setIfFieldExists(payload, fields, ["ErrorTimestamp", "Timestamp"], errorLog.ErrorTimestamp);
  setIfFieldExists(payload, fields, ["UserEmail", "User_x0020_Email"], errorLog.UserEmail);
  setIfFieldExists(payload, fields, ["UserDisplayName", "User_x0020_Display_x0020_Name"], errorLog.UserDisplayName);
  setIfFieldExists(payload, fields, ["SeverityLevel", "Severity_x0020_Level"], errorLog.SeverityLevel);
  setIfFieldExists(payload, fields, ["StackTrace", "Stack_x0020_Trace"], errorLog.StackTrace);
  setIfFieldExists(payload, fields, ["BrowserInfo", "Browser_x0020_Info"], errorLog.BrowserInfo);
  setIfFieldExists(payload, fields, ["UserAction", "User_x0020_Action"], errorLog.UserAction);
  setIfFieldExists(payload, fields, ["AdditionalContext", "Additional_x0020_Context"], errorLog.AdditionalContext);
  setIfFieldExists(payload, fields, ["FeatureFunction", "Feature_x002f_Function"], errorLog.FeatureFunction);
  setIfFieldExists(payload, fields, ["UserRole", "User_x0020_Role"], errorLog.UserRole);

  return payload;
};

const safeJsonStringify = (value: unknown): string => {
  const seen = new WeakSet<object>();
  return JSON.stringify(
    value,
    (_key, item: unknown) => {
      if (typeof item === "object" && item !== null) {
        if (seen.has(item)) {
          return "[Circular]";
        }
        seen.add(item);
      }
      return item;
    },
    2,
  );
};

const normalizeErrorMessage = (reason: unknown): string => {
  if (reason instanceof Error) {
    return reason.message || reason.name || "Error";
  }

  if (typeof reason === "string") {
    return reason;
  }

  if (reason === undefined) {
    return "Unhandled promise rejection with undefined reason";
  }

  if (reason === null) {
    return "Unhandled promise rejection with null reason";
  }

  if (typeof reason === "object") {
    const objectWithMessage = reason as { message?: unknown };
    if (typeof objectWithMessage.message === "string") {
      return objectWithMessage.message;
    }

    try {
      return safeJsonStringify(reason);
    } catch {
      return Object.prototype.toString.call(reason);
    }
  }

  return String(reason);
};

const normalizeStackTrace = (reason: unknown): string => {
  if (reason instanceof Error) {
    return reason.stack || "";
  }

  if (typeof reason === "object" && reason !== null) {
    const objectWithStack = reason as { stack?: unknown };
    if (typeof objectWithStack.stack === "string") {
      return objectWithStack.stack;
    }
  }

  return "";
};

/**
 * Get browser and OS information
 */
export const getBrowserInfo = (): string => {
  const ua = navigator.userAgent;
  let browserName = "Unknown";
  let browserVersion = "Unknown";
  let osName = "Unknown";

  // Detect browser
  if (ua.indexOf("Firefox") > -1) {
    browserName = "Firefox";
    const match = ua.match(/Firefox\/(\d+\.\d+)/);
    browserVersion = match ? match[1] : "Unknown";
  } else if (ua.indexOf("Chrome") > -1) {
    browserName = "Chrome";
    const match = ua.match(/Chrome\/(\d+\.\d+)/);
    browserVersion = match ? match[1] : "Unknown";
  } else if (ua.indexOf("Safari") > -1) {
    browserName = "Safari";
    const match = ua.match(/Version\/(\d+\.\d+)/);
    browserVersion = match ? match[1] : "Unknown";
  } else if (ua.indexOf("Edge") > -1) {
    browserName = "Edge";
    const match = ua.match(/Edge\/(\d+\.\d+)/);
    browserVersion = match ? match[1] : "Unknown";
  } else if (ua.indexOf("Trident") > -1) {
    browserName = "Internet Explorer";
    const match = ua.match(/rv:(\d+\.\d+)/);
    browserVersion = match ? match[1] : "Unknown";
  }

  // Detect OS
  if (ua.indexOf("Win") > -1) {
    osName = "Windows";
  } else if (ua.indexOf("Mac") > -1) {
    osName = "MacOS";
  } else if (ua.indexOf("X11") > -1) {
    osName = "Linux";
  } else if (ua.indexOf("Android") > -1) {
    osName = "Android";
  } else if (ua.indexOf("iPhone") > -1 || ua.indexOf("iPad") > -1) {
    osName = "iOS";
  }

  return `${browserName} ${browserVersion} on ${osName}`;
};

/**
 * Get current user information from SharePoint context
 * Note: Returns "Unknown" as default since we don't have direct access to context in this service.
 * Pass user info via context parameter in captureAndLogError when available.
 */
export const getCurrentUserInfo = async (): Promise<{ email: string; displayName: string }> => {
  try {
    const webUrl = getWebRootUrl();
    const payload = (await fetchJson(
      `${webUrl}/_api/web/currentuser?$select=Email,Title`,
    )) as {
      Email?: string;
      Title?: string;
      d?: {
        Email?: string;
        Title?: string;
      };
    };

    const email = payload.Email || payload.d?.Email || "Unknown";
    const displayName = payload.Title || payload.d?.Title || "Unknown User";

    return {
      email,
      displayName,
    };
  } catch {
    return {
      email: "Unknown",
      displayName: "Unknown User",
    };
  }
};



/**
 * Determine error severity level based on error message and type
 */
export const determineSeverityLevel = (
  errorMessage: string,
  errorType: string
): string => {
  if (
    errorMessage.toLowerCase().includes("critical") ||
    errorMessage.toLowerCase().includes("fatal")
  ) {
    return "Critical";
  }

  if (
    errorType === "Permission Error" ||
    errorType === "Authentication Error"
  ) {
    return "High";
  }

  if (errorType === "Timeout" || errorType === "Network Error") {
    return "High";
  }

  if (errorType === "Validation Error") {
    return "Low";
  }

  return "Medium";
};

/**
 * Log error to SharePoint list asynchronously
 */
export const logErrorToSharePoint = async (
  errorLog: IErrorLog
): Promise<IErrorLogResult> => {
  try {
    const webUrl = getWebRootUrl();
    // Set severity level if not provided
    if (!errorLog.SeverityLevel) {
      errorLog.SeverityLevel = determineSeverityLevel(
        errorLog.ErrorMessage,
        errorLog.ErrorType
      );
    }

    // Ensure timestamp is set
    if (!errorLog.ErrorTimestamp) {
      errorLog.ErrorTimestamp = new Date().toISOString();
    }

    // Get browser info if not provided
    if (!errorLog.BrowserInfo) {
      errorLog.BrowserInfo = getBrowserInfo();
    }



    // Generate title if not provided
    if (!errorLog.Title) {
      errorLog.Title = `Error-${new Date().getTime()}`;
    }

    const resolvedList = await resolveErrorLogList();
    const payload = buildListPayload(errorLog, resolvedList.fields);
    const digest = await getRequestDigest(webUrl);
    const addResult = await fetchJson(
      `${webUrl}/_api/web/${resolvedList.listApi}/items`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify(payload),
      },
    );

    return {
      success: true,
      itemId: extractItemId(addResult),
      message: `Error logged successfully to ${resolvedList.listTitle}`,
    };
  } catch (loggingError) {
    const errorMessage =
      loggingError instanceof Error ? loggingError.message : "Unknown logging error";
    console.error("Error in logErrorToSharePoint:", loggingError);
    console.error("Failed to log:", errorLog);
    return {
      success: false,
      message: errorMessage,
    };
  }
};

/**
 * Capture a caught error and log it to SharePoint
 */
export const captureAndLogError = async (
  error: Error | string,
  context?: {
    component?: string;
    userAction?: string;
    additionalContext?: Record<string, unknown>;
    errorType?: string;
    featureFunction?: string;
    userRole?: string;
    userEmail?: string;
    userDisplayName?: string;
  }
): Promise<IErrorLogResult> => {
  try {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const stackTrace = error instanceof Error ? error.stack || "" : "";
    const userInfo = await getCurrentUserInfo();

    const errorLog: IErrorLog = {
      Title: `Error-${new Date().getTime()}`,
      ErrorMessage: errorMessage,
      ErrorType: context?.errorType || "JavaScript Error",
      ComponentPage: context?.component || "Unknown",
      ErrorTimestamp: new Date().toISOString(),
      UserEmail: context?.userEmail || userInfo.email,
      UserDisplayName: context?.userDisplayName || userInfo.displayName,
      SeverityLevel: determineSeverityLevel(errorMessage, context?.errorType || ""),
      StackTrace: stackTrace,
      BrowserInfo: getBrowserInfo(),
      UserAction: context?.userAction,
      AdditionalContext: context?.additionalContext
        ? JSON.stringify(context.additionalContext)
        : undefined,
      FeatureFunction: context?.featureFunction,
      UserRole: context?.userRole,
    };

    return await logErrorToSharePoint(errorLog);
  } catch (loggingError) {
    console.error("Failed to capture and log error:", loggingError);
    const errorMessage =
      loggingError instanceof Error ? loggingError.message : "Unknown capture error";
    return {
      success: false,
      message: errorMessage,
    };
  }
};

// External CDN domains that host SharePoint platform scripts — errors from these
// are not our code and cannot be fixed by us.
const EXTERNAL_CDN_HOSTS = ["cdn.office.net", "onecdn.static.microsoft"];

// Rate-limit map: errorMessageKey -> last logged timestamp (ms).
// Prevents burst-flooding the log when the same error fires dozens of times
// within a short window (e.g. the Safari chunk-load cascade pattern).
const recentlyLogged = new Map<string, number>();
const RATE_LIMIT_MS = 10_000;

const isRateLimited = (key: string): boolean => {
  const last = recentlyLogged.get(key);
  const now = Date.now();
  if (last !== undefined && now - last < RATE_LIMIT_MS) return true;
  recentlyLogged.set(key, now);
  return false;
};

const shouldIgnoreGlobalError = (event: ErrorEvent): boolean => {
  // Errors originating from SharePoint/Microsoft CDN scripts
  if (EXTERNAL_CDN_HOSTS.some((host) => event.filename?.includes(host))) return true;

  const msg = event.message || "";
  // Benign browser notification — not a real error
  if (msg.includes("ResizeObserver loop")) return true;

  return false;
};

const shouldIgnoreRejection = (event: PromiseRejectionEvent): boolean => {
  const reason = event.reason;

  // undefined/null reasons carry zero diagnostic value and represent ~60% of log noise
  if (reason === undefined || reason === null) return true;

  // Safari wraps unhandled DOM event errors as { isTrusted: true } — not actionable
  if (
    typeof reason === "object" &&
    Object.keys(reason as object).join("") === "isTrusted"
  ) return true;

  if (typeof reason === "object" && reason !== null) {
    const r = reason as Record<string, unknown>;

    // MSAL explicitly marks these as expected — useAuth already logs the auth failure
    if (r.isExpectedFailure === true) return true;

    // SharePoint CDN webpack chunk load failures — infrastructure issue, not our code
    if (r.name === "ChunkLoadError") return true;

    // SharePoint shell iframe request aborted/timed out
    if (r.name === "ShellException") return true;

    // MSAL iframe token acquisition timeout — expected in some auth configurations
    if (r.errorCode === "monitor_window_timeout") return true;
  }

  return false;
};

/**
 * Global error handler for uncaught exceptions
 */
export const setupGlobalErrorHandler = (): void => {
  window.addEventListener("error", (event: ErrorEvent) => {
    if (shouldIgnoreGlobalError(event)) return;

    const message = event.message || normalizeErrorMessage(event.error) || "Unknown error";
    if (isRateLimited(`error:${message}`)) return;

    const globalContext = {
      url: window.location.href,
      source: event.filename,
      line: event.lineno,
      column: event.colno,
      online: navigator.onLine,
      visibility: document.visibilityState,
    };

    void (async () => {
      const userInfo = await getCurrentUserInfo();
      await logErrorToSharePoint({
        Title: `Global Error-${new Date().getTime()}`,
        ErrorMessage: message,
        ErrorType: "JavaScript Error",
        ComponentPage: "Global",
        ErrorTimestamp: new Date().toISOString(),
        UserEmail: userInfo.email,
        UserDisplayName: userInfo.displayName,
        SeverityLevel: "High",
        StackTrace: event.error?.stack || "",
        BrowserInfo: getBrowserInfo(),
        AdditionalContext: safeJsonStringify(globalContext),
      });
    })().catch(() => undefined);
  });

  // Handle unhandled promise rejections
  window.addEventListener("unhandledrejection", (event: PromiseRejectionEvent) => {
    if (shouldIgnoreRejection(event)) return;

    const message = normalizeErrorMessage(event.reason) || "Unhandled Promise Rejection";
    if (isRateLimited(`rejection:${message}`)) return;

    const rejectionContext = {
      url: window.location.href,
      online: navigator.onLine,
      visibility: document.visibilityState,
      reasonType:
        event.reason === null ? "null" : event.reason === undefined ? "undefined" : typeof event.reason,
      reasonRaw:
        typeof event.reason === "object" && event.reason !== null
          ? safeJsonStringify(event.reason)
          : String(event.reason),
    };

    void (async () => {
      const userInfo = await getCurrentUserInfo();
      await logErrorToSharePoint({
        Title: `Promise Rejection-${new Date().getTime()}`,
        ErrorMessage: message,
        ErrorType: "Unhandled Promise Rejection",
        ComponentPage: "Global",
        ErrorTimestamp: new Date().toISOString(),
        UserEmail: userInfo.email,
        UserDisplayName: userInfo.displayName,
        SeverityLevel: "High",
        StackTrace: normalizeStackTrace(event.reason),
        BrowserInfo: getBrowserInfo(),
        AdditionalContext: safeJsonStringify(rejectionContext),
      });
    })().catch(() => undefined);
  });
};
