import { useState, useCallback } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDocument } from "../interfaces/IDocument";

export type ArchiveReplaceStep = "idle" | "confirm" | "archiving" | "replace";

interface IUseArchiveReplaceResult {
  step: ArchiveReplaceStep;
  targetDocument: IDocument | null;
  isArchiving: boolean;
  anchorPosition?: { top: number; left: number };

  // Actions
  initiateArchiveReplace: (
    doc: IDocument,
    anchorPosition?: { top: number; left: number },
  ) => void;
  confirmArchive: () => Promise<boolean>;
  cancelConfirm: () => void;
  cancelReplace: () => void;
  reset: () => void;
}

/**
 * Orchestrates the Archive & Replace flow:
 *
 * 1. initiateArchiveReplace(doc) → opens confirmation modal (step = 'confirm')
 * 2. confirmArchive() → archives the document (step = 'archiving' → 'replace')
 * 3. Replace panel opens with pre-populated metadata
 * 4. On submit (handled by UploadPanel) → new document created
 * 5. reset() → returns to idle
 *
 * If user cancels at step 1 → cancelConfirm() → idle
 * If user cancels at step 3 → cancelReplace() → idle (archive stands)
 */
/**
 * Orchestrates the Archive & Replace flow:
 *
 * 1. initiateArchiveReplace(doc) → opens confirmation modal (step = 'confirm')
 * 2. confirmArchive() → stores the document for archiving on successful upload (step = 'replace')
 * 3. Replace panel opens with pre-populated metadata
 * 4. On submit → DocumentService.uploadDocument() called with archiveTargetId parameter
 * 5. On upload success → DocumentService archives the target document
 * 6. reset() → returns to idle
 *
 * If user cancels at step 1 → cancelConfirm() → idle
 * If user cancels at step 3 → cancelReplace() → idle (no archive happens)
 */
export const useArchiveReplace = (
  context: WebPartContext,
  documentLibrary?: string,
  useMock: boolean = false,
): IUseArchiveReplaceResult => {
  const [step, setStep] = useState<ArchiveReplaceStep>(
    "idle" as ArchiveReplaceStep,
  );
  const [targetDocument, setTargetDocument] = useState<IDocument | null>(null);
  const [isArchiving, setIsArchiving] = useState(false);
  const [anchorPosition, setAnchorPosition] = useState<{
    top: number;
    left: number;
  } | undefined>(undefined);

  const initiateArchiveReplace = useCallback(
    (doc: IDocument, nextAnchorPosition?: { top: number; left: number }) => {
    setTargetDocument(doc);
    setAnchorPosition(nextAnchorPosition);
    setStep("confirm" as ArchiveReplaceStep);
    },
    [],
  );

  const confirmArchive = useCallback(async (): Promise<boolean> => {
    if (!targetDocument) return false;

    // Store the document for later archiving on successful upload
    // Do NOT archive here — archiving is deferred until upload succeeds
    setStep("replace" as ArchiveReplaceStep);
    return true;
  }, [targetDocument]);

  const cancelConfirm = useCallback(() => {
    setStep("idle" as ArchiveReplaceStep);
    setTargetDocument(null);
    setAnchorPosition(undefined);
  }, []);

  const cancelReplace = useCallback(() => {
    // Archive stands — user chose not to upload a replacement
    setStep("idle" as ArchiveReplaceStep);
    setTargetDocument(null);
    setAnchorPosition(undefined);
  }, []);

  const reset = useCallback(() => {
    setStep("idle" as ArchiveReplaceStep);
    setTargetDocument(null);
    setIsArchiving(false);
    setAnchorPosition(undefined);
  }, []);

  return {
    step,
    targetDocument,
    isArchiving,
    anchorPosition,
    initiateArchiveReplace,
    confirmArchive,
    cancelConfirm,
    cancelReplace,
    reset,
  };
};
