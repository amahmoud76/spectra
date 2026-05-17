import {
  ACCEPTED_FILE_EXTENSIONS,
  ACCEPTED_FILE_TYPES,
  MAX_FILE_SIZE_BYTES,
  MAX_FILE_SIZE_DISPLAY,
} from "../config/config";

export interface IFileValidationResult {
  isValid: boolean;
  errorMessage: string;
}

/**
 * Validate a file before upload.
 * Checks: file type, file extension, file size.
 */
export const validateFile = (file: File): IFileValidationResult => {
  // Check file extension
  const extension = "." + (file.name.split(".").pop()?.toLowerCase() || "");
  if (!ACCEPTED_FILE_EXTENSIONS.includes(extension)) {
    return {
      isValid: false,
      errorMessage: `Unsupported file type. Accepted formats: PDF, Word, Excel, PowerPoint.`,
    };
  }

  // Check MIME type (if available — some browsers may not provide it)
  if (file.type && !ACCEPTED_FILE_TYPES.includes(file.type)) {
    return {
      isValid: false,
      errorMessage: `Unsupported file type. Accepted formats: PDF, Word, Excel, PowerPoint.`,
    };
  }

  // Check file size
  if (file.size > MAX_FILE_SIZE_BYTES) {
    return {
      isValid: false,
      errorMessage: `File size exceeds the maximum of ${MAX_FILE_SIZE_DISPLAY}. Please upload a smaller file.`,
    };
  }

  return { isValid: true, errorMessage: "" };
};

/**
 * Format file size for display.
 */
export const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
};

/** Table display without a search match badge (E/C). */
export const FILE_NAME_DISPLAY_MAX_LENGTH = 42;

/** Shorter limit when an E/C badge sits beside the name in the same cell. */
export const FILE_NAME_DISPLAY_MAX_LENGTH_WITH_BADGE = 32;

/**
 * Shorten a file name for table display; full name should be shown in a tooltip.
 */
export const truncateFileNameForDisplay = (
  fileName: string,
  maxLength: number = FILE_NAME_DISPLAY_MAX_LENGTH,
): string => {
  if (fileName.length <= maxLength) {
    return fileName;
  }

  return `${fileName.slice(0, maxLength)}...`;
};
