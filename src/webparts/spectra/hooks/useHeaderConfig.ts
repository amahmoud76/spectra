import { useEffect, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  FALLBACK_HEADER_CONFIG,
  IHeaderConfig,
} from "../interfaces/IHeaderConfig";
import { HeaderConfigService } from "../services/HeaderConfigService";

interface IUseHeaderConfigResult {
  config: IHeaderConfig;
  isLoading: boolean;
  isError: boolean;
}

export const useHeaderConfig = (
  context: WebPartContext,
  useMock: boolean,
): IUseHeaderConfigResult => {
  const [config, setConfig] = useState<IHeaderConfig>(FALLBACK_HEADER_CONFIG);
  const [isLoading, setIsLoading] = useState(true);
  const [isError, setIsError] = useState(false);

  useEffect(() => {
    let cancelled = false;

    const fetchHeaderConfig = async (): Promise<void> => {
      try {
        const service = new HeaderConfigService(context, useMock);
        const result = await service.getConfig();

        if (!cancelled) {
          setConfig(result);
          setIsLoading(false);
        }
      } catch (error) {
        if (!cancelled) {
          console.error("useHeaderConfig: Failed to load header config", error);
          setConfig(FALLBACK_HEADER_CONFIG);
          setIsError(true);
          setIsLoading(false);
        }
      }
    };

    // eslint-disable-next-line no-void
    void fetchHeaderConfig();

    return () => {
      cancelled = true;
    };
  }, [context, useMock]);

  return { config, isLoading, isError };
};
