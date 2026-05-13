export type AuditAction =
  | "view"
  | "upload"
  | "edit"
  | "archive"
  | "replace"
  | "delete";

export interface IAuditEntry {
  id: string; // Unique entry ID
  documentId: string; // Related document ID
  action: AuditAction; // What happened
  performedBy: string; // User email / UPN
  performedByName: string; // User display name
  performedAt: string; // ISO date string
  details?: string; // Optional free-text note
}
