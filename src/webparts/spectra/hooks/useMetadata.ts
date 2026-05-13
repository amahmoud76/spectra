import { useState, useEffect } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  IMetadataOptions,
  EMPTY_METADATA,
} from "../interfaces/IMetadataOptions";
import { MetadataService } from "../services/MetadataService";

interface IUseMetadataResult {
  options: IMetadataOptions;
  isLoading: boolean;
  isError: boolean;
}

/**
 * Load all metadata options from reference lists.
 * Results are cached by MetadataService for 24 hours.
 * Returns EMPTY_METADATA while loading — safe for components
 * to render immediately without null checks.
 */
export const useMetadata = (
  context: WebPartContext,
  useMock: boolean,
): IUseMetadataResult => {
  const [options, setOptions] = useState<IMetadataOptions>(EMPTY_METADATA);
  const [isLoading, setIsLoading] = useState(true);
  const [isError, setIsError] = useState(false);

  useEffect(() => {
    let cancelled = false;

    const fetchMetadata = async (): Promise<void> => {
      try {
        const service = new MetadataService(context, useMock);
        const result = await service.getOptions();

        if (!cancelled) {
          setOptions(result);
          setIsLoading(false);
        }
      } catch (error) {
        if (!cancelled) {
          console.error("useMetadata: Failed to load options", error);
          setIsError(true);
          setIsLoading(false);
        }
      }
    };

    void fetchMetadata();

    return () => {
      cancelled = true;
    };
  }, [context, useMock]);

  return { options, isLoading, isError };
};
