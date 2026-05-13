import { IAuthResponse } from "../interfaces/IAuthResponse";
import { AUTH_CACHE_KEY_PREFIX, AUTH_CACHE_TTL_MS } from "../config/config";
import { getTime } from "date-fns";

/**
 * Read cached auth response for a user from localStorage.
 * Returns null if not found or expired.
 */
export const getCachedAuth = (userEmail: string): IAuthResponse | null => {
  try {
    const key = AUTH_CACHE_KEY_PREFIX + userEmail;
    const raw = localStorage.getItem(key);
    if (!raw) return null;

    const { data, expiry } = JSON.parse(raw);
    if (getTime(new Date()) > expiry) {
      localStorage.removeItem(key);
      return null;
    }

    return data as IAuthResponse;
  } catch {
    return null;
  }
};

/**
 * Write auth response to localStorage cache.
 * Cached for 1 hour (AUTH_CACHE_TTL_MS).
 */
export const setCachedAuth = (
  userEmail: string,
  response: IAuthResponse,
): void => {
  try {
    const key = AUTH_CACHE_KEY_PREFIX + userEmail;
    localStorage.setItem(
      key,
      JSON.stringify({
        data: response,
        expiry: getTime(new Date()) + AUTH_CACHE_TTL_MS,
      }),
    );
  } catch {
    // localStorage quota exceeded — proceed without caching
  }
};
