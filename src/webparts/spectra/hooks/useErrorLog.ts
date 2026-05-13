import { useCallback } from "react";
import {
  captureAndLogError,
  IErrorLogResult,
} from "../services/errorLogService";

interface IErrorContext {
  component?: string;
  userAction?: string;
  additionalContext?: Record<string, unknown>;
  errorType?: string;
  featureFunction?: string;
}

interface IErrorLogReturn {
  logError: (error: Error | string, context?: IErrorContext) => Promise<IErrorLogResult>;
}

/**
 * Hook to easily log errors within components
 * @returns logError - Function to call when an error occurs
 */
export const useErrorLog = (): IErrorLogReturn => {
  const logError = useCallback(
    async (error: Error | string, context?: IErrorContext) => {
      return captureAndLogError(error, context);
    },
    []
  );

  return { logError };
};
