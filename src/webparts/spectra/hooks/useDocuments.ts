import { useState, useEffect, useCallback, useRef } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDocument } from "../interfaces/IDocument";
import { DocumentService } from "../services/DocumentService";
import { EffectiveRole } from "../interfaces/IAuthResponse";
import { PAGE_SIZE_DEFAULT } from "../config/config";

interface IUseDocumentsResult {
  documents: IDocument[];
  totalCount: number;
  isLoading: boolean;
  isError: boolean;
  refetch: () => void;
}

/**
 * Fetch documents from the document library.
 *
 * - Admin sees all documents (Current + Archive)
 * - Contributor and Viewer see Current only
 * - Documents are NOT fetched on initial load (search-first home page)
 * - Call refetch() after a search or filter is applied
 */
export const useDocuments = (
  context: WebPartContext,
  role: EffectiveRole,
  pageSize: number = PAGE_SIZE_DEFAULT,
  documentLibrary?: string,
  includeArchivedForAdmin: boolean = false,
  useMock: boolean = false,
): IUseDocumentsResult => {
  const [documents, setDocuments] = useState([] as IDocument[]);
  const [totalCount, setTotalCount] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [isError, setIsError] = useState(false);
  const [fetchTrigger, setFetchTrigger] = useState(0);
  const hasSearched = useRef(false);

  const refetch = useCallback(() => {
    hasSearched.current = true;
    setFetchTrigger((n) => n + 1);
  }, []);

  useEffect(() => {
    // Don't fetch on initial load — search-first home page
    if (!hasSearched.current) return;

    let cancelled = false;

    const fetchDocuments = async (): Promise<void> => {
      setIsLoading(true);
      setIsError(false);

      try {
        const service = new DocumentService(context, documentLibrary, useMock);
        const statusFilter =
          role === "admin"
            ? includeArchivedForAdmin
              ? "All"
              : "Current"
            : "Current";
        const result = await service.getDocuments(statusFilter, 1, pageSize);

        if (!cancelled) {
          setDocuments(result.documents);
          setTotalCount(result.totalCount);
          setIsLoading(false);
        }
      } catch (error) {
        if (!cancelled) {
          console.error("useDocuments: Failed to fetch", error);
          setIsError(true);
          setIsLoading(false);
        }
      }
    };

    void fetchDocuments();

    return () => {
      cancelled = true;
    };
  }, [fetchTrigger, role, pageSize, context, documentLibrary, includeArchivedForAdmin, useMock]);

  return { documents, totalCount, isLoading, isError, refetch };
};
