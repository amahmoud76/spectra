import * as React from "react";
import { IWebPartProps } from "../interfaces/IWebPartProps";
import { IDocument } from "../interfaces/IDocument";
import { IUploadPayload } from "../interfaces/IUploadPayload";
import { IFilterState, defaultFilterState } from "../interfaces/IFilterState";
import { IMetadataOption } from "../interfaces/IMetadataOptions";

// Hooks
import { useAuth } from "../hooks/useAuth";
import { useDocuments } from "../hooks/useDocuments";
import { useMetadata } from "../hooks/useMetadata";
import { useFilters } from "../hooks/useFilters";
import { useSorting } from "../hooks/useSorting";
import { usePagination } from "../hooks/usePagination";
import { useNotification } from "../hooks/useNotification";
import { useUpload } from "../hooks/useUpload";
import { useArchiveReplace } from "../hooks/useArchiveReplace";
import { useHeaderConfig } from "../hooks/useHeaderConfig";

// Utilities
import {
  applyFiltersWithMeta,
  hasActiveSearchText,
} from "../utils/filterHelper";
import { exportDocumentsToCSV } from "../utils/exportHelper";
import { generateFileName } from "../utils/fileNamingHelper";
import {
  AUTH_CACHE_KEY_PREFIX,
  METADATA_CACHE_KEY,
  HEADER_CONFIG_CACHE_KEY,
  GENERIC_ERROR_MESSAGE,
} from "../config/config";

// Services
import { DocumentService } from "../services/DocumentService";
import { setupGlobalErrorHandler } from "../services/errorLogService";

// Components — Core
import { ErrorBoundary } from "./ErrorBoundary/ErrorBoundary";
import { ParentHeader } from "../components/ParentHeader/ParentHeader";
import { Footer } from "../components/Footer/Footer";
import { SearchBar } from "../components/SearchBar/SearchBar";
import { Toolbar } from "../components/Toolbar/Toolbar";
import { DataTable } from "../components/DataTable/DataTable";
import { Pagination } from "../components/Pagination/Pagination";

// Components — Panels
import { FilterPanel } from "../components/FilterPanel/FilterPanel";
import { UploadPanel } from "../components/UploadPanel/UploadPanel";
import { EditPanel } from "../components/EditPanel/EditPanel";
import { ConfirmDialog } from "../components/ConfirmDialog/ConfirmDialog";
import { SuccessBanner } from "../components/SuccessBanner/SuccessBanner";
import { ErrorBanner } from "../components/ErrorBanner/ErrorBanner";
import { AuthWarningStrip } from "../components/AuthWarningStrip/AuthWarningStrip";
import { EmptyState } from "../components/EmptyState/EmptyState";
import { SplashScreen } from "../components/SplashScreen/SplashScreen";

// Components — Pages
import { DocumentViewingPage } from "../components/DocumentViewingPage/DocumentViewingPage";
import { ShowArchivedToggle } from "./ShowArchivedToggle/ShowArchivedToggle";
import { ViewFullLibraryButton } from "./ViewFullLibraryButton/ViewFullLibraryButton";
import { ActiveFilterChips } from "./ActiveFilterChips/ActiveFilterChips";
// Styles
import styles from "./SPECTRA.module.scss";

// Types
type PageState = "landing" | "results" | "viewing";

const hasValidCacheEntry = (cacheKey: string): boolean => {
  try {
    const raw = localStorage.getItem(cacheKey);
    if (!raw) return false;

    const parsed: { expiry?: number } = JSON.parse(raw);
    return typeof parsed.expiry === "number" && Date.now() <= parsed.expiry;
  } catch {
    return false;
  }
};

export const SPECTRA: React.FC<IWebPartProps> = ({
  context,
  pageSize,
  title,
  userEmail,
  enableDevRoleSwitch,
  documentLibrary,
  inactivityTimeoutMinutes,
  useMock,
  useAdGroups,
  mockRole,
  enableStartupSplash,
  enableVerboseStartupStatus,
  startupSplashCompletionDelayMs,
  initialDocumentId,
}) => {
  // ── Page state ──────────────────────────────────────────────
  const [page, setPage] = React.useState<PageState>("landing");
  const [searchText, setSearchText] = React.useState("");
  const [showArchivedDocuments, setShowArchivedDocuments] =
    React.useState(false);
  const [isFullLibraryView, setIsFullLibraryView] = React.useState(false);
  const [hasResultsContext, setHasResultsContext] = React.useState(false);
  const [viewingDocument, setViewingDocument] =
    React.useState<IDocument | null>(null);

  // ── Panel visibility ────────────────────────────────────────
  const [filterPanelOpen, setFilterPanelOpen] = React.useState(false);
  const [draftFilters, setDraftFilters] =
    React.useState<IFilterState>(defaultFilterState);
  const [filterPanelResetToken, setFilterPanelResetToken] = React.useState(0);
  const [uploadPanelOpen, setUploadPanelOpen] = React.useState(false);
  const [uploadPanelSession, setUploadPanelSession] = React.useState(0);
  const [editPanelOpen, setEditPanelOpen] = React.useState(false);
  const [editTarget, setEditTarget] = React.useState<IDocument | null>(null);

  // ── Delete confirmation ─────────────────────────────────────
  const [deleteTarget, setDeleteTarget] = React.useState<IDocument | null>(
    null,
  );

  // ── Archive confirmation (admin direct archive) ─────────────
  const [archiveTarget, setArchiveTarget] = React.useState<IDocument | null>(
    null,
  );

  // ── Re-activate confirmation ────────────────────────────────
  const [reActivateTarget, setReActivateTarget] =
    React.useState<IDocument | null>(null);
  const [reActivateNamingChoice, setReActivateNamingChoice] = React.useState<{
    doc: IDocument;
    newFileName: string;
    fromEditPanel: boolean;
  } | null>(null);
  const [isSplashFading, setIsSplashFading] = React.useState(false);

  // ── Deep-link state ─────────────────────────────────────────
  const [deepLinkLoading, setDeepLinkLoading] =
    React.useState(!!initialDocumentId);
  const [deepLinkError, setDeepLinkError] = React.useState(false);
  const deepLinkHandled = React.useRef(false);

  const [showStartupSplash, setShowStartupSplash] = React.useState(() => {
    if (!enableStartupSplash || useMock) return false;

    const hasAuthCache = hasValidCacheEntry(AUTH_CACHE_KEY_PREFIX + userEmail);
    const hasMetadataCache = hasValidCacheEntry(METADATA_CACHE_KEY);
    const hasHeaderCache = hasValidCacheEntry(HEADER_CONFIG_CACHE_KEY);

    return !(hasAuthCache && hasMetadataCache && hasHeaderCache);
  });

  // ── Duplicate confirmation ──────────────────────────────────
  const [duplicateConfirmation, setDuplicateConfirmation] = React.useState<{
    isOpen: boolean;
    existingDocId: string;
    existingFileName: string;
    pendingPayload: IUploadPayload | null;
    isReplace: boolean;
  }>({
    isOpen: false,
    existingDocId: "",
    existingFileName: "",
    pendingPayload: null,
    isReplace: false,
  });

  // ── Edit duplicate confirmation ────────────────────────────
  const [editDuplicateConfirmation, setEditDuplicateConfirmation] =
    React.useState<{
      isOpen: boolean;
      editingDocId: string;
      existingDocId: string;
      existingFileName: string;
      pendingUpdates: Partial<IUploadPayload> | null;
    }>({
      isOpen: false,
      editingDocId: "",
      existingDocId: "",
      existingFileName: "",
      pendingUpdates: null,
    });
  // ── Hooks ───────────────────────────────────────────────────
  const auth = useAuth(
    context,
    userEmail,
    enableDevRoleSwitch,
    useMock,
    useAdGroups,
    mockRole,
  );
  const metadata = useMetadata(context, useMock);
  const documents = useDocuments(
    context,
    auth.effectiveRole,
    pageSize,
    documentLibrary,
    showArchivedDocuments,
    useMock,
  );
  const filters = useFilters();
  const sorting = useSorting();
  const notification = useNotification();
  const upload = useUpload(context, documentLibrary, useMock);
  const archiveReplace = useArchiveReplace(context, documentLibrary, useMock);
  const headerConfig = useHeaderConfig(context, useMock);

  // ── Derived data ────────────────────────────────────────────
  // Helper: Create a map from metadata option value to searchTokens
  const tokenMapFromOptions = (
    options: IMetadataOption[],
  ): Map<string, string[]> => {
    const map = new Map<string, string[]>();
    options.forEach((option) => {
      map.set(option.value, option.searchTokens);
    });
    return map;
  };

  // Memoized token maps for enrichment lookup
  const metadataTokenMaps = React.useMemo(
    () => ({
      documentTypeMap: tokenMapFromOptions(metadata.options.documentTypes),
      assetMap: tokenMapFromOptions(metadata.options.assets),
      therapeuticAreaMap: tokenMapFromOptions(
        metadata.options.therapeuticAreas,
      ),
      indicationMap: tokenMapFromOptions(metadata.options.indications),
    }),
    [
      metadata.options.documentTypes,
      metadata.options.assets,
      metadata.options.therapeuticAreas,
      metadata.options.indications,
    ],
  );

  // Memoized enriched documents with merged search tokens
  const enrichedDocuments = React.useMemo(() => {
    return documents.documents.map((doc) => ({
      ...doc,
      searchTokens: [
        ...doc.searchTokens,
        ...(metadataTokenMaps.documentTypeMap.get(doc.documentType) || []),
        ...doc.asset.flatMap((a) => metadataTokenMaps.assetMap.get(a) || []),
        ...doc.therapeuticArea.flatMap(
          (ta) => metadataTokenMaps.therapeuticAreaMap.get(ta) || [],
        ),
        ...doc.indication.flatMap(
          (i) => metadataTokenMaps.indicationMap.get(i) || [],
        ),
      ],
    }));
  }, [documents.documents, metadataTokenMaps]);

  const filteredResult = applyFiltersWithMeta(
    enrichedDocuments,
    filters.filters,
  );
  const filteredDocuments = filteredResult.documents;

  const sortedDocuments = React.useMemo(() => {
    const sorted = sorting.sortDocs(filteredDocuments);
    // Only group exact-before-close when the user hasn't explicitly chosen a column sort.
    // If the user clicked a sort column, respect that order instead.
    if (!hasActiveSearchText(filters.filters) || sorting.userHasExplicitSort) {
      return sorted;
    }

    const exact: IDocument[] = [];
    const close: IDocument[] = [];
    sorted.forEach((doc) => {
      if (filteredResult.matchKindByDocumentId.get(doc.id) === "close") {
        close.push(doc);
      } else {
        exact.push(doc);
      }
    });
    return [...exact, ...close];
  }, [
    filteredDocuments,
    sorting.sortState,
    sorting.userHasExplicitSort,
    filters.filters,
    filteredResult.matchKindByDocumentId,
  ]);

  const hasSearchApplied = hasActiveSearchText(filters.filters);

  const useEnhancedStyle = true;

  const pagination = usePagination(sortedDocuments, pageSize);

  // const noResults = !documents.isLoading && sortedDocuments.length === 0;
  const noResults = !documents.isLoading && sortedDocuments.length === 0;

  // const hasPanelFilters = filters.activeFilterCount > 0;
  const hasPanelFilters = filters.activeFilterCount > 0;

  // const showFilteredEmptyState =
  //   noResults && hasPanelFilters && !hasSearchApplied;

  // Show "No results found" when filters are applied (with or without search)
  const showFilteredEmptyState = noResults && hasPanelFilters;

  // const showSearchNoResults = noResults && hasSearchApplied;
  // Show "No document found" only when search is applied WITHOUT any panel filters
  const showSearchNoResults = noResults && hasSearchApplied && !hasPanelFilters;
  const showResultsIdleEmptyState =
    page === "results" &&
    hasResultsContext &&
    !hasSearchApplied &&
    !hasPanelFilters &&
    !isFullLibraryView;

  const siteUrl = context.pageContext.web.absoluteUrl;
  const userDisplayName = context.pageContext.user.displayName;
  const repositoryTitle = title || "SPECTRA Document Repository";

  React.useEffect(() => {
    if (!enableStartupSplash || useMock) {
      setShowStartupSplash(false);
      setIsSplashFading(false);
      return;
    }

    const hasAuthCache = hasValidCacheEntry(AUTH_CACHE_KEY_PREFIX + userEmail);
    const hasMetadataCache = hasValidCacheEntry(METADATA_CACHE_KEY);
    const hasHeaderCache = hasValidCacheEntry(HEADER_CONFIG_CACHE_KEY);

    const shouldShow = !(hasAuthCache && hasMetadataCache && hasHeaderCache);
    setShowStartupSplash(shouldShow);
    if (shouldShow) {
      setIsSplashFading(false);
    }
  }, [enableStartupSplash, useMock, userEmail]);

  React.useEffect(() => {
    if (!showStartupSplash) return;

    if (!auth.isLoading && !metadata.isLoading && !headerConfig.isLoading) {
      setIsSplashFading(true);
      const hideTimer = window.setTimeout(() => {
        setShowStartupSplash(false);
        setIsSplashFading(false);
      }, startupSplashCompletionDelayMs);

      return () => window.clearTimeout(hideTimer);
    }
  }, [
    showStartupSplash,
    auth.isLoading,
    metadata.isLoading,
    headerConfig.isLoading,
    startupSplashCompletionDelayMs,
  ]);

  // ── Deep-link: open document from ?spectraDoc= URL param ───
  React.useEffect(() => {
    if (!initialDocumentId || auth.isLoading || deepLinkHandled.current) return;
    deepLinkHandled.current = true;

    const service = new DocumentService(context, documentLibrary, useMock);
    service
      .getDocumentById(String(initialDocumentId))
      .then((doc) => {
        if (doc) {
          setViewingDocument(doc);
          setPage("viewing");
        } else {
          setDeepLinkError(true);
        }
        setDeepLinkLoading(false);
      })
      .catch(() => {
        setDeepLinkError(true);
        setDeepLinkLoading(false);
      });
  }, [auth.isLoading, initialDocumentId, context, documentLibrary, useMock]);

  // ── Setup global error handler ──────────────────────────────
  React.useEffect(() => {
    setupGlobalErrorHandler();
  }, []);

  const startupProgress = React.useMemo(() => {
    const wrapLabel = (detailedLabel: string): string =>
      enableVerboseStartupStatus ? detailedLabel : "Loading SPECTRA";

    if (auth.isLoading) {
      if (auth.startupStage === "authenticating") {
        return { label: wrapLabel("Connecting to SharePoint"), percent: 20 };
      }
      if (auth.startupStage === "loadingUserInfo") {
        return { label: wrapLabel("Loading user information"), percent: 38 };
      }
      if (auth.startupStage === "assigningRole") {
        return { label: wrapLabel("Assigning role"), percent: 55 };
      }
      if (auth.startupStage === "retrying") {
        return {
          label: wrapLabel("Connection is slow — trying again…"),
          percent: 25,
        };
      }

      return { label: wrapLabel("Connecting to SharePoint"), percent: 20 };
    }

    if (metadata.isLoading) {
      return { label: wrapLabel("Pulling metadata"), percent: 75 };
    }

    if (headerConfig.isLoading) {
      return { label: wrapLabel("Building interface"), percent: 90 };
    }

    return {
      label: enableVerboseStartupStatus ? "Ready" : "Loading complete",
      percent: 100,
    };
  }, [
    auth.isLoading,
    auth.startupStage,
    metadata.isLoading,
    headerConfig.isLoading,
    enableVerboseStartupStatus,
  ]);

  const handleViewFullLibrary = React.useCallback(() => {
    setShowArchivedDocuments(false);
    setIsFullLibraryView(true);
    setHasResultsContext(true);
    setPage("results");
    documents.refetch();
  }, [documents]);

  const handleArchiveToggleChange = React.useCallback(
    (next: boolean) => {
      setShowArchivedDocuments(next);
      if (isFullLibraryView) {
        setHasResultsContext(true);
        setPage("results");
      }
      documents.refetch();
    },
    [documents, isFullLibraryView],
  );

  // ── Search handler ──────────────────────────────────────────
  const handleSearch = React.useCallback(() => {
    filters.setFilter("searchText", searchText);
    if (searchText || filters.hasActiveFilters) {
      documents.refetch();
      setHasResultsContext(true);
      setPage("results");
    }
  }, [searchText, filters, documents]);

  const handleClearSearch = React.useCallback(() => {
    setSearchText("");
    filters.setFilter("searchText", "");
    setHasResultsContext(true);
    setPage("results");
  }, [filters]);

  const handleSearchTextChange = React.useCallback(
    (value: string) => {
      setSearchText(value);

      if (value) {
        return;
      }

      if (!filters.filters.searchText) {
        return;
      }

      filters.setFilter("searchText", "");
      setHasResultsContext(true);
      setPage("results");
    },
    [filters],
  );

  // ── Filter apply ────────────────────────────────────────────
  const handleFilterApply = React.useCallback(() => {
    // filters.setFilter("searchText", draftFilters.searchText);
    filters.setFilter("documentType", draftFilters.documentType);
    filters.setFilter("therapeuticArea", draftFilters.therapeuticArea);
    filters.setFilter("subTherapeuticArea", draftFilters.subTherapeuticArea);
    filters.setFilter("asset", draftFilters.asset);
    filters.setFilter("indication", draftFilters.indication);
    filters.setFilter("lineOfTherapy", draftFilters.lineOfTherapy);
    filters.setFilter("paid", draftFilters.paid);
    filters.setFilter("diseaseArea", draftFilters.diseaseArea);
    filters.setFilter("effectiveDateFrom", draftFilters.effectiveDateFrom);
    filters.setFilter("effectiveDateTo", draftFilters.effectiveDateTo);
    filters.setFilter("uploadDateFrom", draftFilters.uploadDateFrom);
    filters.setFilter("uploadDateTo", draftFilters.uploadDateTo);

    setFilterPanelOpen(false);
    documents.refetch();
    setHasResultsContext(true);
    setPage("results");
  }, [documents, draftFilters, filters]);

  const handleOpenFilterPanel = React.useCallback(() => {
    setDraftFilters(filters.filters);
    setFilterPanelOpen(true);
  }, [filters.filters]);

  const handleFilterCancel = React.useCallback(() => {
    setDraftFilters(filters.filters);
    setFilterPanelOpen(false);
  }, [filters.filters]);

  const handleFilterReset = React.useCallback(() => {
    setDraftFilters(defaultFilterState);
    setFilterPanelResetToken((prev) => prev + 1);
  }, []);

  const handleClearFiltersQuick = React.useCallback(() => {
    // Preserve the search text when clearing filters
    const currentSearchText = filters.filters.searchText;
    filters.clearAllFilters();
    filters.setFilter("searchText", currentSearchText);
    setDraftFilters((_prev) => ({
      ...defaultFilterState,
      searchText: currentSearchText,
    }));
    setFilterPanelResetToken((prev) => prev + 1);
    documents.refetch();
    setHasResultsContext(true);
  }, [filters, documents]);

  const handleDraftFilterChange = React.useCallback(
    <K extends keyof IFilterState>(field: K, value: IFilterState[K]) => {
      setDraftFilters((prev) => ({ ...prev, [field]: value }));
    },
    [],
  );

  const handleRemoveFilterValue = React.useCallback(
    (field: keyof IFilterState, value: string) => {
      const current = filters.filters[field] as string[];
      const updated = current.filter((v) => v !== value);
      filters.setFilter(field, updated as IFilterState[typeof field]);
      documents.refetch();
    },
    [filters, documents],
  );

  const handleRemoveDateFilter = React.useCallback(
    (type: "effective" | "upload") => {
      if (type === "effective") {
        filters.setFilter("effectiveDateFrom", null);
        filters.setFilter("effectiveDateTo", null);
      } else {
        filters.setFilter("uploadDateFrom", null);
        filters.setFilter("uploadDateTo", null);
      }
      documents.refetch();
    },
    [filters, documents],
  );

  // ── Document click → viewing page ───────────────────────────
  const handleDocumentClick = React.useCallback((doc: IDocument) => {
    setViewingDocument(doc);
    setPage("viewing");
  }, []);

  // ── Back to results ─────────────────────────────────────────
  const handleBackToResults = React.useCallback(() => {
    setViewingDocument(null);
    setPage("results");
  }, []);

  // ── Export CSV (Admin only) ────────────────────────────────
  const handleExport = React.useCallback(async () => {
    try {
      const service = new DocumentService(context, documentLibrary, useMock);
      const result = await service.getDocuments("All", 1, pageSize);
      exportDocumentsToCSV(result.documents);
    } catch {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [context, documentLibrary, useMock, pageSize, notification]);

  // ── Logo click → landing page ───────────────────────────────
  const handleLogoClick = React.useCallback(() => {
    setViewingDocument(null);
    setPage("landing");
    setHasResultsContext(false);
    setIsFullLibraryView(false);
    setShowArchivedDocuments(false);
    setSearchText("");
    filters.clearAllFilters();
  }, [filters]);

  const handleOpenUploadPanel = React.useCallback(() => {
    setUploadPanelSession((prev) => prev + 1);
    setUploadPanelOpen(true);
  }, []);

  // ── Upload submit ───────────────────────────────────────────
  const handleUploadSubmit = React.useCallback(
    async (payload: IUploadPayload) => {
      const result = await upload.upload(payload);

      // Check for duplicate identity
      if (result.isDuplicateIdentity && result.existingDocumentId) {
        notification.clearNotification();
        setDuplicateConfirmation({
          isOpen: true,
          existingDocId: result.existingDocumentId,
          existingFileName: result.existingFileName || "",
          pendingPayload: payload,
          isReplace: false,
        });
        return;
      }

      if (result.success) {
        setUploadPanelOpen(false);
        upload.resetUpload();
        notification.showSuccess(
          "File successfully uploaded.",
          result.generatedFileName,
        );
        documents.refetch();
      } else {
        notification.showError(GENERIC_ERROR_MESSAGE);
      }
    },
    [upload, notification, documents],
  );

  // ── Edit save ───────────────────────────────────────────────
  const handleEditSave = React.useCallback(
    async (documentId: string, updates: Partial<IUploadPayload>) => {
      const service = new DocumentService(context, documentLibrary, useMock);
      const success = await service.updateDocument(documentId, updates);
      if (success) {
        setEditPanelOpen(false);
        setEditTarget(null);
        notification.showSuccess("Document updated successfully.");
        documents.refetch();
      } else {
        notification.showError(GENERIC_ERROR_MESSAGE);
      }
    },
    [context, documentLibrary, useMock, notification, documents, page],
  );

  // ── Edit duplicate check ────────────────────────────────────
  const handleEditDuplicateCheck = React.useCallback(
    async (documentId: string, updates: Partial<IUploadPayload>) => {
      const service = new DocumentService(context, documentLibrary, useMock);
      const duplicateCheck = await service.checkEditDuplicate(
        documentId,
        updates,
      );

      if (duplicateCheck.isDuplicate && duplicateCheck.duplicateId) {
        // Duplicate found — show confirmation dialog
        notification.clearNotification();
        setEditDuplicateConfirmation({
          isOpen: true,
          editingDocId: documentId,
          existingDocId: duplicateCheck.duplicateId,
          existingFileName: duplicateCheck.duplicateFileName || "",
          pendingUpdates: updates,
        });
        return;
      }

      // No duplicate — proceed with edit
      await handleEditSave(documentId, updates);
    },
    [context, documentLibrary, useMock, handleEditSave, notification],
  );

  // ── Delete confirm ──────────────────────────────────────────
  const handleDeleteConfirm = React.useCallback(async () => {
    if (!deleteTarget) return;
    if (auth.effectiveRole !== "admin") {
      notification.showError(GENERIC_ERROR_MESSAGE);
      return;
    }
    const service = new DocumentService(context, documentLibrary, useMock);
    const success = await service.deleteDocument(deleteTarget.id);
    if (success) {
      setDeleteTarget(null);
      notification.showSuccess(
        `"${deleteTarget.fileName}" deleted successfully.`,
      );
      documents.refetch();
      if (page === "viewing") setPage("results");
    } else {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [
    deleteTarget,
    auth.effectiveRole,
    context,
    documentLibrary,
    useMock,
    notification,
    documents,
    page,
  ]);

  // ── Admin direct archive confirm ────────────────────────────
  const handleArchiveConfirm = React.useCallback(async () => {
    if (!archiveTarget) return;
    const service = new DocumentService(context, documentLibrary, useMock);
    const success = await service.archiveDocument(archiveTarget.id);
    if (success) {
      setArchiveTarget(null);
      notification.showSuccess(
        `"${archiveTarget.fileName}" archived successfully.`,
      );
      documents.refetch();
    } else {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [
    archiveTarget,
    context,
    documentLibrary,
    useMock,
    notification,
    documents,
  ]);

  // ── Re-activate confirm ─────────────────────────────────────
  // Shared execution — called after the naming choice is resolved
  const doReActivate = React.useCallback(
    async (
      doc: IDocument,
      newFileName: string | undefined,
      fromEditPanel: boolean,
    ) => {
      const service = new DocumentService(context, documentLibrary, useMock);
      const success = await service.reActivateDocument(doc.id, newFileName);
      if (success) {
        setReActivateTarget(null);
        setReActivateNamingChoice(null);
        if (fromEditPanel) {
          setEditPanelOpen(false);
          setEditTarget(null);
        }
        const displayName = newFileName ?? doc.fileName;
        notification.showSuccess(`"${displayName}" re-activated successfully.`);
        documents.refetch();
      } else {
        notification.showError(GENERIC_ERROR_MESSAGE);
      }
    },
    [context, documentLibrary, useMock, notification, documents],
  );

  // Naming check — computes new name and routes to the right modal
  const triggerReActivateWithNamingCheck = React.useCallback(
    (doc: IDocument, fromEditPanel: boolean) => {
      const newFileName = generateFileName(doc, doc.fileExtension);
      if (newFileName.toLowerCase() !== doc.fileName.toLowerCase()) {
        setReActivateNamingChoice({ doc, newFileName, fromEditPanel });
      } else if (fromEditPanel) {
        void doReActivate(doc, undefined, true);
      } else {
        setReActivateTarget(doc);
      }
    },
    [doReActivate],
  );

  const handleReActivateConfirm = React.useCallback(async () => {
    if (!reActivateTarget) return;
    await doReActivate(reActivateTarget, undefined, false);
  }, [reActivateTarget, doReActivate]);

  // ── Duplicate confirmation dialog ──────────────────────────
  const handleDuplicateConfirmArchive = React.useCallback(async () => {
    if (!duplicateConfirmation.pendingPayload) return;

    // Re-upload with archiveTargetId
    const result = duplicateConfirmation.isReplace
      ? await upload.upload(
          duplicateConfirmation.pendingPayload,
          duplicateConfirmation.existingDocId,
        )
      : await upload.upload(
          duplicateConfirmation.pendingPayload,
          duplicateConfirmation.existingDocId,
        );

    setDuplicateConfirmation({
      isOpen: false,
      existingDocId: "",
      existingFileName: "",
      pendingPayload: null,
      isReplace: false,
    });

    if (result.success) {
      setUploadPanelOpen(false);
      upload.resetUpload();
      notification.showSuccess(
        duplicateConfirmation.isReplace
          ? "File has been archived and replaced."
          : "File uploaded and existing document archived.",
        result.generatedFileName,
      );
      documents.refetch();
      if (page === "viewing") setPage("results");
    } else if (!result.isDuplicateIdentity) {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [duplicateConfirmation, upload, notification, documents, page]);

  const handleDuplicateConfirmView = React.useCallback(async () => {
    notification.clearNotification();
    const existingDocId = duplicateConfirmation.existingDocId;

    setDuplicateConfirmation({
      isOpen: false,
      existingDocId: "",
      existingFileName: "",
      pendingPayload: null,
      isReplace: false,
    });
    setUploadPanelOpen(false);

    if (!existingDocId) {
      return;
    }

    let doc = documents.documents.find((d) => d.id === existingDocId) ?? null;

    if (!doc) {
      const service = new DocumentService(context, documentLibrary, useMock);
      doc = await service.getDocumentById(existingDocId);
    }

    if (doc) {
      setViewingDocument(doc);
      setPage("viewing");
    } else {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [
    duplicateConfirmation.existingDocId,
    documents.documents,
    context,
    documentLibrary,
    useMock,
    notification,
  ]);

  const handleDuplicateConfirmCancel = React.useCallback(() => {
    notification.clearNotification();
    setDuplicateConfirmation({
      isOpen: false,
      existingDocId: "",
      existingFileName: "",
      pendingPayload: null,
      isReplace: false,
    });
  }, [notification]);

  // ── Edit duplicate confirmation dialog ──────────────────────
  const handleEditDuplicateConfirmArchive = React.useCallback(async () => {
    if (!editDuplicateConfirmation.pendingUpdates) return;

    // Archive the duplicate document first
    const service = new DocumentService(context, documentLibrary, useMock);
    const archiveSuccess = await service.archiveDocument(
      editDuplicateConfirmation.existingDocId,
    );

    if (archiveSuccess) {
      // Proceed with the edit
      const success = await service.updateDocument(
        editDuplicateConfirmation.editingDocId,
        editDuplicateConfirmation.pendingUpdates,
      );

      if (success) {
        setEditPanelOpen(false);
        setEditTarget(null);
        notification.showSuccess(
          "Document updated and existing duplicate archived.",
        );
        documents.refetch();
      } else {
        notification.showError(GENERIC_ERROR_MESSAGE);
      }
    } else {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }

    setEditDuplicateConfirmation({
      isOpen: false,
      editingDocId: "",
      existingDocId: "",
      existingFileName: "",
      pendingUpdates: null,
    });
  }, [
    editDuplicateConfirmation,
    context,
    documentLibrary,
    useMock,
    notification,
    documents,
  ]);

  const handleEditDuplicateConfirmCancel = React.useCallback(() => {
    notification.clearNotification();
    setEditDuplicateConfirmation({
      isOpen: false,
      editingDocId: "",
      existingDocId: "",
      existingFileName: "",
      pendingUpdates: null,
    });
  }, [notification]);

  const handleEditDuplicateConfirmView = React.useCallback(async () => {
    notification.clearNotification();
    const existingDocId = editDuplicateConfirmation.existingDocId;

    setEditDuplicateConfirmation({
      isOpen: false,
      editingDocId: "",
      existingDocId: "",
      existingFileName: "",
      pendingUpdates: null,
    });
    setEditPanelOpen(false);

    if (!existingDocId) {
      return;
    }

    let doc = documents.documents.find((d) => d.id === existingDocId) ?? null;

    if (!doc) {
      const service = new DocumentService(context, documentLibrary, useMock);
      doc = await service.getDocumentById(existingDocId);
    }

    if (doc) {
      setEditTarget(null);
      setViewingDocument(doc);
      setPage("viewing");
    } else {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [
    editDuplicateConfirmation.existingDocId,
    documents.documents,
    context,
    documentLibrary,
    useMock,
    notification,
  ]);
  // ── Archive & Replace: confirm step ─────────────────────────
  const handleArchiveReplaceConfirm = React.useCallback(async () => {
    const success = await archiveReplace.confirmArchive();
    if (!success) {
      notification.showError(GENERIC_ERROR_MESSAGE);
    }
  }, [archiveReplace, notification]);

  // ── Archive & Replace: replacement submit ───────────────────
  const handleReplaceSubmit = React.useCallback(
    async (payload: IUploadPayload) => {
      const result = await upload.upload(
        payload,
        archiveReplace.targetDocument?.id,
      );

      // Check for duplicate identity (in replace flow)
      if (result.isDuplicateIdentity && result.existingDocumentId) {
        notification.clearNotification();
        setDuplicateConfirmation({
          isOpen: true,
          existingDocId: result.existingDocumentId,
          existingFileName: result.existingFileName || "",
          pendingPayload: payload,
          isReplace: true,
        });
        return;
      }

      if (result.success) {
        archiveReplace.reset();
        upload.resetUpload();
        notification.showSuccess(
          "File has been archived and replaced.",
          result.generatedFileName,
        );
        documents.refetch();
        if (page === "viewing") setPage("results");
      } else {
        notification.showError(GENERIC_ERROR_MESSAGE);
      }
    },
    [upload, archiveReplace, notification, documents, page],
  );

  // ── Build pre-populated data for Replace panel ──────────────
  const replacePrePopData = React.useMemo(():
    | Partial<IUploadPayload>
    | undefined => {
    const doc = archiveReplace.targetDocument;
    if (!doc) return undefined;
    return {
      documentType: doc.documentType,
      therapeuticArea: doc.therapeuticArea,
      subTherapeuticArea: doc.subTherapeuticArea,
      asset: doc.asset,
      indication: doc.indication,
      lineOfTherapy: doc.lineOfTherapy,
      paid: doc.paid,
      diseaseArea: doc.diseaseArea,
      effectiveDate: doc.effectiveDate,
      description: doc.description,
      comments: "",
    };
  }, [archiveReplace.targetDocument]);

  const handleReActivate = React.useCallback(() => {
    if (!editTarget) return;
    triggerReActivateWithNamingCheck(editTarget, true);
  }, [editTarget, triggerReActivateWithNamingCheck]);

  // ── Startup splash (cold start / expired cache) ────────────
  if (showStartupSplash) {
    return (
      <div className={styles.spectraApp}>
        <SplashScreen
          appName="SPECTRA"
          subtitle="Strategic Portfolio Enterprise Content Tracking Repository App"
          statusLabel={startupProgress.label}
          progressPercent={startupProgress.percent}
          isFading={isSplashFading}
        />
      </div>
    );
  }

  // ── Loading state ───────────────────────────────────────────
  if (auth.isLoading) {
    return (
      <div className={styles.spectraApp}>
        <ParentHeader
          config={headerConfig.config}
          role="viewer"
          userDisplayName={userDisplayName}
          userEmail={userEmail}
          siteUrl={siteUrl}
          enableDevRoleSwitch={false}
          onRoleBadgeClick={() => undefined}
          onSpectraClick={() => undefined}
        />
        <div style={{ padding: 40, textAlign: "center" }}>
          <span className={styles.spinner} />
          <p>Loading...</p>
        </div>
      </div>
    );
  }

  // ── Deep-link loading ───────────────────────────────────────
  if (deepLinkLoading) {
    return (
      <div className={styles.spectraApp}>
        <ParentHeader
          config={headerConfig.config}
          role="viewer"
          userDisplayName={userDisplayName}
          userEmail={userEmail}
          siteUrl={siteUrl}
          enableDevRoleSwitch={false}
          onRoleBadgeClick={() => undefined}
          onSpectraClick={() => undefined}
        />
        <div style={{ padding: 40, textAlign: "center" }}>
          <span className={styles.spinner} />
          <p>Loading document…</p>
        </div>
      </div>
    );
  }

  // ── Deep-link error ─────────────────────────────────────────
  if (deepLinkError) {
    return (
      <div className={styles.spectraApp}>
        <ParentHeader
          config={headerConfig.config}
          role={auth.effectiveRole}
          userDisplayName={userDisplayName}
          userEmail={userEmail}
          siteUrl={siteUrl}
          enableDevRoleSwitch={enableDevRoleSwitch}
          onRoleBadgeClick={() => undefined}
          onSpectraClick={() => {
            setDeepLinkError(false);
            setPage("landing");
          }}
        />
        <div style={{ padding: 40, textAlign: "center" }}>
          <p>
            The document you are trying to reach is currently unavailable. It
            may have been deleted or moved. Our system refreshes nightly, so
            please try again later.
          </p>
          <button
            onClick={() => {
              setDeepLinkError(false);
              setPage("landing");
            }}
          >
            Go to SPECTRA Search
          </button>
        </div>
      </div>
    );
  }

  // ── Render ──────────────────────────────────────────────────
  return (
    <ErrorBoundary>
      <div className={styles.spectraApp}>
        <ParentHeader
          config={headerConfig.config}
          role={auth.effectiveRole}
          userDisplayName={userDisplayName}
          userEmail={userEmail}
          siteUrl={siteUrl}
          enableDevRoleSwitch={auth.isDevRoleSwitchEnabled}
          onRoleBadgeClick={auth.cycleDevRole}
          onSpectraClick={handleLogoClick}
        />

        {/* Notification banners */}
        {auth.isError && (
          <AuthWarningStrip onRetry={auth.retryAuth} userEmail={userEmail} />
        )}
        {notification.notification?.type === "success" && (
          <SuccessBanner
            message={notification.notification.message}
            fileName={notification.notification.fileName}
            onDismiss={notification.clearNotification}
          />
        )}
        {notification.notification?.type === "error" && (
          <ErrorBanner
            message={notification.notification.message}
            onDismiss={notification.clearNotification}
          />
        )}

        {/* ── LANDING PAGE ─────────────────────────────────────── */}
        {page === "landing" && (
          <div className={styles.landingPage}>
            <div className={styles.pageHeaderRow}>
              <h1 className={styles.pageTitle}>{repositoryTitle}</h1>
              <Toolbar
                role={auth.effectiveRole}
                activeFilterCount={filters.activeFilterCount}
                onFilterClick={handleOpenFilterPanel}
                onClearFiltersClick={handleClearFiltersQuick}
                onUploadClick={handleOpenUploadPanel}
                showUpload={true}
                showFilter={true}
                showExport={false}
              />
            </div>

            <div className={styles.pageSearchRow}>
              <SearchBar
                className={styles.searchGrow}
                value={searchText}
                onChange={handleSearchTextChange}
                onSearch={handleSearch}
                onClear={handleClearSearch}
                isError={noResults && !!searchText}
              />
              {auth.effectiveRole === "admin" && (
                <div className={styles.adminSearchControls}>
                  <ViewFullLibraryButton
                    onClick={handleViewFullLibrary}
                    isActive={isFullLibraryView}
                  />
                </div>
              )}
            </div>

            <div className={styles.landingWelcome}>
              <h2 className={styles.landingWelcomeTitle}>
                <img
                  src={require("../assets/icons/rainbow-half.svg")}
                  alt=""
                  className={styles.spectraRainbowIcon}
                  style={{ display: "inline", width: "1em", height: "1em" }}
                  aria-hidden="true"
                />
                Welcome to SPECTRA
              </h2>
              <p className={styles.landingWelcomeSubline}>
                Strategic Portfolio Enterprise Content Tracking Repository App
              </p>
              <p>
                SPECTRA is AbbVie&apos;s centralized repository for strategic
                documents.
              </p>
            </div>

            <div className={styles.landingSecondaryCard}>
              <div>
                <p>
                  SPECTRA, AbbVie&apos;s centralized repository for strategic
                  documents, provides a single, reliable location to organize,
                  search, and access key portfolio materials.
                </p>
                <p>
                  SPECTRA supports efficient document retrieval, consistent
                  classification, and streamlined management of current content
                  across teams.
                </p>
                <p>
                  SPECTRA includes the following approved document types
                  (additional document types to be added in future releases.)
                </p>
                <ul>
                  <li>Disease Area Strategies (DAS)</li>
                  <li>Target Product Profiles (TPP)</li>
                  <li>Target Product Claims (TPC)</li>
                  <li>Integrated Evidence Plans (IEP)</li>
                  <li>Early Brand Plans (EBP)</li>
                  <li>Early Integrated Value Propositions (EIVP)</li>
                  <li>Integrated Access Strategies (IAS)</li>
                </ul>
              </div>
              <div className={styles.landingInfoPanel}>
                <h3>
                  <img
                    src={require("../assets/icons/rainbow-half.svg")}
                    alt=""
                    className={styles.spectraRainbowIcon}
                    style={{ display: "inline", width: "1em", height: "1em" }}
                    aria-hidden="true"
                  />
                  How to Use SPECTRA
                </h3>
                <p>
                  Use the search and filter tools to quickly locate documents by
                  name, asset, document type, or therapeutic area.
                </p>
                <p>
                  SPECTRA is intended to serve as a trusted source for current
                  strategic documents and to support effective review,
                  reporting, and portfolio decision-making.
                </p>
              </div>
            </div>
          </div>
        )}

        {/* ── SEARCH RESULTS PAGE ──────────────────────────────── */}
        {page === "results" && (
          <div className={styles.resultsPage}>
            <div className={styles.pageHeaderRow}>
              <div className={styles.pageTitleGroup}>
                <button
                  className={styles.homeIconBtn}
                  onClick={handleLogoClick}
                  aria-label="Back to SPECTRA home"
                  type="button"
                >
                  <img
                    src={require("../assets/icons/home.svg")}
                    alt=""
                    style={{
                      width: "16px",
                      height: "16px",
                      display: "inline-block",
                    }}
                    aria-hidden="true"
                  />
                </button>
                <h1 className={styles.pageTitle}>{repositoryTitle}</h1>
              </div>
              <Toolbar
                role={auth.effectiveRole}
                activeFilterCount={filters.activeFilterCount}
                onUploadClick={handleOpenUploadPanel}
                onExportClick={handleExport}
                showUpload={true}
                showFilter={false}
                showClearFilters={false}
                showExport={auth.effectiveRole === "admin"}
              />
            </div>

            <div className={styles.pageSearchRow}>
              <SearchBar
                className={styles.searchGrow}
                value={searchText}
                onChange={handleSearchTextChange}
                onSearch={handleSearch}
                onClear={handleClearSearch}
                isError={noResults && !!searchText}
              />

              {useEnhancedStyle && hasSearchApplied && (
                <div className={styles.exactMatchLegend}>
                  <img
                    src={require("../assets/icons/check-circle.svg")}
                    alt=""
                    style={{ width: 17, height: 16, display: "block" }}
                    aria-hidden="true"
                  />
                  = Exact match
                </div>
              )}

              {auth.effectiveRole === "admin" && (
                <div className={styles.adminSearchControls}>
                  <ShowArchivedToggle
                    checked={showArchivedDocuments}
                    onChange={handleArchiveToggleChange}
                  />
                  <ViewFullLibraryButton
                    onClick={handleViewFullLibrary}
                    isActive={isFullLibraryView}
                  />
                </div>
              )}

              <Toolbar
                role={auth.effectiveRole}
                activeFilterCount={filters.activeFilterCount}
                onFilterClick={handleOpenFilterPanel}
                onClearFiltersClick={handleClearFiltersQuick}
                showUpload={false}
                showFilter={true}
                showExport={false}
              />
            </div>

            {filters.activeFilterCount > 0 && (
              <ActiveFilterChips
                filters={filters.filters}
                onRemoveValue={handleRemoveFilterValue}
                onRemoveDateRange={handleRemoveDateFilter}
                onClearAll={handleClearFiltersQuick}
              />
            )}

            {isFullLibraryView && (
              <div
                className={styles.noResultsHint}
                role="status"
                aria-live="polite"
              >
                {showArchivedDocuments
                  ? "Showing full document library including archived documents."
                  : "Showing full document library (active documents only). Use the archive toggle to include archived documents."}
              </div>
            )}

            {documents.isLoading ? (
              <DataTable
                documents={[]}
                role={auth.effectiveRole}
                sortState={sorting.sortState}
                onSort={sorting.toggleSort}
                onDocumentClick={handleDocumentClick}
                onArchiveReplaceClick={(doc, anchorPosition) =>
                  archiveReplace.initiateArchiveReplace(doc, anchorPosition)
                }
                isLoading={true}
                useEnhancedStyle={useEnhancedStyle}
              />
            ) : showResultsIdleEmptyState ? (
              <EmptyState type="results-idle" />
            ) : showFilteredEmptyState ? (
              <EmptyState type="no-results" />
            ) : showSearchNoResults ? (
              <>
                <div className={styles.noResultsHint}>
                  <img
                    src={require("../assets/icons/circle-exclamation.svg")}
                    alt=""
                    style={{
                      width: "16px",
                      height: "16px",
                      display: "inline-block",
                      marginRight: "6px",
                    }}
                    aria-hidden="true"
                  />
                  No document found
                </div>
                <DataTable
                  documents={[]}
                  role={auth.effectiveRole}
                  sortState={sorting.sortState}
                  onSort={sorting.toggleSort}
                  onDocumentClick={handleDocumentClick}
                  onEditClick={(doc) => {
                    setEditTarget(doc);
                    setEditPanelOpen(true);
                  }}
                  onArchiveClick={(doc) => setArchiveTarget(doc)}
                  onDeleteClick={(doc) => setDeleteTarget(doc)}
                  onArchiveReplaceClick={(doc, anchorPosition) =>
                    archiveReplace.initiateArchiveReplace(doc, anchorPosition)
                  }
                  onReActivateClick={(doc) =>
                    triggerReActivateWithNamingCheck(doc, false)
                  }
                  isLoading={false}
                  useEnhancedStyle={useEnhancedStyle}
                />
              </>
            ) : (
              <>
                <DataTable
                  documents={pagination.paginatedDocuments}
                  role={auth.effectiveRole}
                  sortState={sorting.sortState}
                  onSort={sorting.toggleSort}
                  onDocumentClick={handleDocumentClick}
                  searchMatchKindByDocumentId={
                    filteredResult.matchKindByDocumentId
                  }
                  onEditClick={(doc) => {
                    setEditTarget(doc);
                    setEditPanelOpen(true);
                  }}
                  onArchiveClick={(doc) => setArchiveTarget(doc)}
                  onDeleteClick={(doc) => setDeleteTarget(doc)}
                  onArchiveReplaceClick={(doc, anchorPosition) =>
                    archiveReplace.initiateArchiveReplace(doc, anchorPosition)
                  }
                  onReActivateClick={(doc) =>
                    triggerReActivateWithNamingCheck(doc, false)
                  }
                  isLoading={false}
                  useEnhancedStyle={useEnhancedStyle}
                />
                <Pagination
                  currentPage={pagination.currentPage}
                  totalPages={pagination.totalPages}
                  pageSize={pagination.pageSize}
                  pageSizeOptions={pagination.pageSizeOptions}
                  startIndex={pagination.startIndex}
                  endIndex={pagination.endIndex}
                  totalCount={sortedDocuments.length}
                  canGoNext={pagination.canGoNext}
                  canGoPrevious={pagination.canGoPrevious}
                  onPageChange={pagination.goToPage}
                  onPageSizeChange={pagination.setPageSize}
                  onFirst={pagination.goToFirstPage}
                  onPrevious={pagination.goToPreviousPage}
                  onNext={pagination.goToNextPage}
                  onLast={pagination.goToLastPage}
                />
              </>
            )}
          </div>
        )}

        {/* ── DOCUMENT VIEWING PAGE ────────────────────────────── */}
        {page === "viewing" && viewingDocument && (
          <DocumentViewingPage
            document={viewingDocument}
            role={auth.effectiveRole}
            siteUrl={siteUrl}
            onHome={handleLogoClick}
            onBack={handleBackToResults}
            onEditClick={() => {
              setEditTarget(viewingDocument);
              setEditPanelOpen(true);
            }}
            onArchiveClick={() => setArchiveTarget(viewingDocument)}
            onDeleteClick={() => setDeleteTarget(viewingDocument)}
            onArchiveReplaceClick={() =>
              archiveReplace.initiateArchiveReplace(viewingDocument)
            }
          />
        )}

        {/* ── PANELS ───────────────────────────────────────────── */}

        {/* Upload Panel */}
        <UploadPanel
          key={`upload-panel-${uploadPanelSession}`}
          isOpen={uploadPanelOpen}
          options={metadata.options}
          isAdmin={auth.effectiveRole === "admin"}
          isContributor={auth.effectiveRole === "contributor"}
          isUploading={upload.isUploading}
          uploadProgress={upload.uploadProgress}
          uploadPercent={upload.uploadPercent}
          cancelUpload={upload.cancelUpload}
          onSubmit={handleUploadSubmit}
          onCancel={() => {
            setUploadPanelOpen(false);
            upload.resetUpload();
          }}
        />

        {/* Filter Panel */}
        <FilterPanel
          isOpen={filterPanelOpen}
          filters={draftFilters}
          resetToken={filterPanelResetToken}
          options={metadata.options}
          onFilterChange={handleDraftFilterChange}
          onApply={handleFilterApply}
          onCancel={handleFilterCancel}
          onReset={handleFilterReset}
        />

        {/* Edit Panel (Admin only) */}
        <EditPanel
          isOpen={editPanelOpen}
          isAdmin={auth.effectiveRole === "admin"}
          document={editTarget}
          options={metadata.options}
          isSaving={false}
          onSave={handleEditDuplicateCheck}
          onCancel={() => {
            setEditPanelOpen(false);
            setEditTarget(null);
          }}
          onArchiveClick={() => {
            setEditPanelOpen(false);
            if (editTarget) setArchiveTarget(editTarget);
          }}
          onDeleteClick={() => {
            setEditPanelOpen(false);
            if (editTarget) setDeleteTarget(editTarget);
          }}
          onReplaceFileClick={() => {
            setEditPanelOpen(false);
            if (editTarget) archiveReplace.initiateArchiveReplace(editTarget);
          }}
          onReActivateClick={handleReActivate}
        />

        {/* Archive & Replace: Replace Panel */}
        {archiveReplace.step === "replace" && (
          <UploadPanel
            isOpen={true}
            options={metadata.options}
            isAdmin={auth.effectiveRole === "admin"}
            isContributor={auth.effectiveRole === "contributor"}
            isUploading={upload.isUploading}
            uploadProgress={upload.uploadProgress}
            uploadPercent={upload.uploadPercent}
            cancelUpload={upload.cancelUpload}
            panelNotice="The current file will be archived after successful replacement upload."
            onSubmit={handleReplaceSubmit}
            onCancel={() => archiveReplace.cancelReplace()}
            title="Replace"
            prePopulatedData={replacePrePopData}
            archiveTargetDocument={archiveReplace.targetDocument || undefined}
          />
        )}

        {/* ── CONFIRMATION DIALOGS ─────────────────────────────── */}

        {/* Duplicate Detection: Confirmation */}
        <ConfirmDialog
          isOpen={duplicateConfirmation.isOpen}
          title="An active document already exists for this metadata"
          message={`"${duplicateConfirmation.existingFileName}" already matches this metadata. You can archive it and continue, view it, or cancel.`}
          confirmLabel="Archive Active and Proceed"
          onConfirm={handleDuplicateConfirmArchive}
          onCancel={handleDuplicateConfirmCancel}
          secondaryLabel="View Existing"
          onSecondary={handleDuplicateConfirmView}
        />

        {/* Edit Duplicate Detection: Confirmation */}
        <ConfirmDialog
          isOpen={editDuplicateConfirmation.isOpen}
          title="This metadata change would create a duplicate"
          message={`"${editDuplicateConfirmation.existingFileName}" already matches the metadata you entered. You can archive it and continue, view it, or cancel.`}
          confirmLabel="Archive Existing and Proceed"
          onConfirm={handleEditDuplicateConfirmArchive}
          onCancel={handleEditDuplicateConfirmCancel}
          secondaryLabel="View Existing"
          onSecondary={handleEditDuplicateConfirmView}
        />

        {/* Archive & Replace: Confirmation */}
        <ConfirmDialog
          isOpen={
            archiveReplace.step === "confirm" ||
            archiveReplace.step === "archiving"
          }
          variant="compact"
          anchorPosition={archiveReplace.anchorPosition}
          title="Replace this file?"
          message="After your replacement upload succeeds, the active file will be archived automatically. Metadata stays editable for the new version."
          confirmLabel="Proceed with Replacement"
          onConfirm={handleArchiveReplaceConfirm}
          onCancel={archiveReplace.cancelConfirm}
        />

        {/* Admin Direct Archive: Confirmation */}
        <ConfirmDialog
          isOpen={archiveTarget !== null}
          title="Archive this document?"
          message={`"${archiveTarget?.fileName}" will be archived.`}
          confirmLabel="Archive"
          onConfirm={handleArchiveConfirm}
          onCancel={() => setArchiveTarget(null)}
        />

        {/* Re-activate: Confirmation (no name change needed) */}
        <ConfirmDialog
          isOpen={reActivateTarget !== null}
          title="Re-activate this document?"
          message={`"${reActivateTarget?.fileName}" will be moved back to Active status.`}
          confirmLabel="Re-activate"
          onConfirm={handleReActivateConfirm}
          onCancel={() => setReActivateTarget(null)}
        />

        {/* Re-activate: Naming choice (document name is outdated) */}
        <ConfirmDialog
          isOpen={reActivateNamingChoice !== null}
          title="Re-activate this document?"
          message={
            reActivateNamingChoice && (
              <div>
                <p>
                  The stored document name does not match the current naming
                  convention based on its metadata. Choose how to proceed.
                </p>
                <div style={{ marginTop: 12 }}>
                  <div
                    style={{
                      fontSize: 12,
                      color: "var(--spectra-text-secondary)",
                      marginBottom: 4,
                    }}
                  >
                    Current name
                  </div>
                  <div
                    style={{
                      fontFamily: "Consolas, 'Courier New', monospace",
                      fontSize: 13,
                    }}
                  >
                    {reActivateNamingChoice.doc.fileName}
                  </div>
                </div>
                <div style={{ marginTop: 12 }}>
                  <div
                    style={{
                      fontSize: 12,
                      color: "var(--spectra-text-secondary)",
                      marginBottom: 4,
                    }}
                  >
                    Updated name
                  </div>
                  <div className={styles.namingImmutablePreview}>
                    {reActivateNamingChoice.newFileName}
                  </div>
                </div>
              </div>
            )
          }
          confirmLabel={`Update to new name`}
          secondaryLabel="Keep old name"
          cancelLabel="Cancel"
          onConfirm={() =>
            reActivateNamingChoice &&
            void doReActivate(
              reActivateNamingChoice.doc,
              reActivateNamingChoice.newFileName,
              reActivateNamingChoice.fromEditPanel,
            )
          }
          onSecondary={() =>
            reActivateNamingChoice &&
            void doReActivate(
              reActivateNamingChoice.doc,
              undefined,
              reActivateNamingChoice.fromEditPanel,
            )
          }
          onCancel={() => setReActivateNamingChoice(null)}
        />

        {/* Delete: Confirmation */}
        <ConfirmDialog
          isOpen={deleteTarget !== null}
          title="Delete this document?"
          message={`"${deleteTarget?.fileName}" will be permanently deleted. This action cannot be undone.`}
          confirmLabel="Delete"
          isDestructive={true}
          onConfirm={handleDeleteConfirm}
          onCancel={() => setDeleteTarget(null)}
        />

        <Footer inactivityTimeoutMinutes={inactivityTimeoutMinutes} />
      </div>
    </ErrorBoundary>
  );
};
