import { IDocument } from "./IDocument";

export interface IDocumentResponse {
  documents: IDocument[]; // Documents for the current page
  totalCount: number; // Total count after filtering (for pagination)
  page: number; // Current page number (1-based)
  pageSize: number; // Number of items per page
}
