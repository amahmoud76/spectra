import { useState, useCallback } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DocumentService } from "../services/DocumentService";
import { captureAndLogError } from "../services/errorLogService";
import { IUploadPayload, IUploadResult } from "../interfaces/IUploadPayload";

type UploadProgressStage = "Preparing" | "Uploading" | "Finalizing" | "Completed" | "Failed" | "";

interface IUseUploadResult {
  isUploading: boolean;
  uploadProgress: UploadProgressStage;
  uploadResult: IUploadResult | null;
  upload: (payload: IUploadPayload, archiveTargetId?: string) => Promise<IUploadResult>;
  resetUpload: () => void;
}

export const useUpload = (
  context: WebPartContext,
  documentLibrary?: string,
  useMock: boolean = false,
): IUseUploadResult => {
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<UploadProgressStage>("");
  const [uploadResult, setUploadResult] = useState<IUploadResult | null>(null);

  const upload = useCallback(
    async (payload: IUploadPayload, archiveTargetId?: string): Promise<IUploadResult> => {
      let currentStage: UploadProgressStage = "Preparing";
      setIsUploading(true);
      setUploadProgress("Preparing");
      setUploadResult(null);

      try {
        // Simulate brief preparing stage for UX feedback
        await new Promise(resolve => setTimeout(resolve, 300));
        currentStage = "Uploading";
        setUploadProgress("Uploading");

        const service = new DocumentService(context, documentLibrary, useMock);
        const result = await service.uploadDocument(payload, archiveTargetId);
        
        currentStage = "Finalizing";
        setUploadProgress("Finalizing");
        
        // Simulate brief finalizing stage for UX feedback
        await new Promise(resolve => setTimeout(resolve, 200));
        
        if (result.success) {
          currentStage = "Completed";
          setUploadProgress("Completed");
        } else {
          currentStage = "Failed";
          setUploadProgress("Failed");

          if (!result.isDuplicateIdentity) {
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
        currentStage = "Failed";
        setUploadProgress("Failed");
        const errorMessage =
          error instanceof Error ? error.message : String(error);

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
      }
    },
    [context, documentLibrary, useMock],
  );

  const resetUpload = useCallback(() => {
    setIsUploading(false);
    setUploadProgress("");
    setUploadResult(null);
  }, []);

  return { isUploading, uploadProgress, uploadResult, upload, resetUpload };
};
