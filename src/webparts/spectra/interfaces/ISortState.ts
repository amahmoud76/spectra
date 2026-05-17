// Only single-value columns are sortable
// Multi-value columns (asset, therapeuticArea, subTherapeuticArea,
// indication, paid) are NOT sortable
export type SortField =
  | "fileName"
  | "comments"
  | "documentType"
  | "effectiveDate"
  | "uploadDate"
  | "createdBy"
  | "modifiedBy"
  | "lineOfTherapy"
  | "diseaseArea"
  | "status";

export type SortDirection = "asc" | "desc";

export interface ISortState {
  field: SortField;
  direction: SortDirection;
}
