import { useState, useCallback, useRef, useEffect } from "react";

export type NotificationType = "success" | "error";

interface INotification {
  message: string;
  type: NotificationType;
  fileName?: string;
}

interface IUseNotificationResult {
  notification: INotification | null;
  showSuccess: (message: string, fileName?: string) => void;
  showError: (message: string) => void;
  clearNotification: () => void;
}

export const useNotification = (): IUseNotificationResult => {
  const [notification, setNotification] = useState<INotification | null>(null);
  const timerRef = useRef<number | null>(null);

  // Clear any existing timer
  const clearTimer = useCallback(() => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current);
      timerRef.current = null;
    }
  }, []);

  const clearNotification = useCallback(() => {
    clearTimer();
    setNotification(null);
  }, [clearTimer]);

  const showSuccess = useCallback(
    (message: string, fileName?: string) => {
      clearTimer();
      setNotification({
        message,
        type: "success" as NotificationType,
        fileName,
      });
    },
    [clearTimer],
  );

  const showError = useCallback(
    (message: string) => {
      clearTimer();
      // Error banners persist until manually dismissed — no auto-dismiss
      setNotification({ message, type: "error" as NotificationType });
    },
    [clearTimer],
  );

  // Cleanup timer on unmount
  useEffect(() => {
    return () => clearTimer();
  }, [clearTimer]);

  return { notification, showSuccess, showError, clearNotification };
};
