import { useState, useCallback, useRef } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DocumentService } from "../services/DocumentService";
import { captureAndLogError } from "../services/errorLogService";
import { IUploadPayload, IUploadResult } from "../interfaces/IUploadPayload";

export type UploadProgressStage =
  | "Preparing"
  | "UploadingFile"
  | "SavingMetadata"
  | "Finalizing"
  | "Completed"
  | "Failed"
  | "";

interface IUseUploadResult {
  isUploading: boolean;
  uploadProgress: UploadProgressStage;
  uploadPercent: number;
  uploadResult: IUploadResult | null;
  upload: (payload: IUploadPayload, archiveTargetId?: string) => Promise<IUploadResult>;
  cancelUpload: () => void;
  resetUpload: () => void;
}

export const useUpload = (
  context: WebPartContext,
  documentLibrary?: string,
  useMock: boolean = false,
): IUseUploadResult => {
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<UploadProgressStage>("");
  const [uploadPercent, setUploadPercent] = useState(0);
  const [uploadResult, setUploadResult] = useState<IUploadResult | null>(null);

  // Holds the XHR abort function while a file transfer is in progress.
  const cancelRef = useRef<(() => void) | null>(null);

  const cancelUpload = useCallback(() => {
    cancelRef.current?.();
  }, []);

  const upload = useCallback(
    async (payload: IUploadPayload, archiveTargetId?: string): Promise<IUploadResult> => {
      let currentStage: UploadProgressStage = "Preparing";
      setIsUploading(true);
      setUploadProgress("Preparing");
      setUploadPercent(0);
      setUploadResult(null);

      // cancelSignal is passed into DocumentService so it can wire the XHR abort.
      const cancelSignal = { abort: (): void => { /* filled in by service */ } };
      cancelRef.current = () => cancelSignal.abort();

      try {
        await new Promise((resolve) => setTimeout(resolve, 300));

        currentStage = "UploadingFile";
        setUploadProgress("UploadingFile");

        const service = new DocumentService(context, documentLibrary, useMock);
        const result = await service.uploadDocument(payload, archiveTargetId, {
          onFileProgress: (percent) => {
            setUploadPercent(percent);
          },
          onMetadataSave: () => {
            currentStage = "SavingMetadata";
            setUploadProgress("SavingMetadata");
            setUploadPercent(100);
          },
          cancelSignal,
        });

        currentStage = "Finalizing";
        setUploadProgress("Finalizing");
        await new Promise((resolve) => setTimeout(resolve, 200));

        if (result.success) {
          currentStage = "Completed";
          setUploadProgress("Completed");
        } else {
          currentStage = "Failed";
          setUploadProgress("Failed");

          // Skip logging for expected user-correctable outcomes: duplicate identity
          // (user sees a dialog) and file name conflict (user sees an actionable message).
          if (!result.isDuplicateIdentity && !result.isFileNameConflict) {
            await captureAndLogError(result.message || "Upload failed", {
              component: "Upload Hook",
              errorType: "Upload Failure",
              featureFunction: "Upload",
              userAction: "Submit upload",
              userEmail: context.pageContext.user.email,
              userDisplayName: context.pageContext.user.displayName,
              additionalContext: {
                stage: currentStage,
                documentLibrary: documentLibrary || "default",
                useMock,
                archiveTargetId: archiveTargetId || null,
                resultMessage: result.message,
                fileName: payload.file?.name,
                fileType: payload.file?.type,
                fileSizeBytes: payload.file?.size,
                documentType: payload.documentType,
                therapeuticArea: payload.therapeuticArea,
                asset: payload.asset,
              },
            });
          }
        }

        setUploadResult(result);
        return result;
      } catch (error) {
        // Treat user-initiated cancellation as a silent reset, not an error.
        if ((error as Error & { cancelled?: boolean }).cancelled) {
          setUploadProgress("");
          setUploadPercent(0);
          setIsUploading(false);
          cancelRef.current = null;
          const cancelled: IUploadResult = { success: false, message: "Upload cancelled." };
          setUploadResult(cancelled);
          return cancelled;
        }

        currentStage = "Failed";
        setUploadProgress("Failed");
        const rawMessage =
          error instanceof Error ? error.message : String(error);
        // The browser throws a NotReadableError/DOMException when the file handle
        // is revoked after selection (e.g., user moved or deleted the file).
        const isFileUnreadable =
          /could not be read|NotReadableError|permission.*file|file.*permission/i.test(rawMessage) ||
          (error instanceof DOMException && error.name === "NotReadableError");
        const errorMessage = isFileUnreadable
          ? "The file could not be read. Please re-select the file and try again."
          : rawMessage;

        await captureAndLogError(errorMessage, {
          component: "Upload Hook",
          errorType: "Upload Exception",
          featureFunction: "Upload",
          userAction: "Submit upload",
          userEmail: context.pageContext.user.email,
          userDisplayName: context.pageContext.user.displayName,
          additionalContext: {
            stage: currentStage,
            documentLibrary: documentLibrary || "default",
            useMock,
            archiveTargetId: archiveTargetId || null,
            fileName: payload.file?.name,
            fileType: payload.file?.type,
            fileSizeBytes: payload.file?.size,
            documentType: payload.documentType,
            therapeuticArea: payload.therapeuticArea,
            asset: payload.asset,
          },
        });

        const result: IUploadResult = {
          success: false,
          message: `An error occurred and no file was uploaded. ${errorMessage}`,
        };
        setUploadResult(result);
        return result;
      } finally {
        setIsUploading(false);
        cancelRef.current = null;
      }
    },
    [context, documentLibrary, useMock],
  );

  const resetUpload = useCallback(() => {
    setIsUploading(false);
    setUploadProgress("");
    setUploadPercent(0);
    setUploadResult(null);
  }, []);

  return { isUploading, uploadProgress, uploadPercent, uploadResult, upload, cancelUpload, resetUpload };
};
